/* global Word, console */

/**
 * ResolveSuggestionCommand — Command pattern for accepting/rejecting suggestions.
 *
 * Encapsulates the complete logic for resolving a single `Suggestion` in Word:
 * finding the Content Control, observing tracked changes through multiple
 * evidence sources (CC-scoped, CC-range, body-level, operational anchor,
 * colocated comment range), applying the terminal action, and cleaning up
 * artifacts.
 *
 * Parallel to `ApplySuggestionCommand` which handles *applying* suggestions.
 * This command handles *resolving* them after the user accepts or rejects.
 *
 * Never throws — catches all errors and returns a `SuggestionActionResult`.
 *
 * @module ResolveSuggestionCommand
 */

import type { IResolutionObservabilityPort } from "../../domain/ports";
import { ExecuteResolutionStateMachine } from "../../domain/suggestion/ExecuteResolutionStateMachine";
import type {
  ResolutionExecutionReport,
  Suggestion,
  SuggestionActionResult,
} from "../../domain/types";
import { NoopResolutionObservabilityAdapter } from "../observability/NoopResolutionObservabilityAdapter";
import { CommentOnlySuggestionResolver } from "./resolution/CommentOnlySuggestionResolver";
import { DocumentReviewStateInspector } from "./resolution/DocumentReviewStateInspector";
import { ReplaceResolutionWorkflow } from "./resolution/ReplaceResolutionOrchestrator";
import type { ReplaceResolutionStrategy } from "./resolution/ReplaceResolutionStrategyContext";
import type { ResolutionObservation } from "./resolution/ResolutionContext";
import { ResolutionErrorSerializer } from "./resolution/ResolutionErrorParser";
import { ResolutionObservabilityReporter } from "./resolution/ResolutionObservabilityAdapter";
import { describeTrackedChangesForLog } from "./resolution/ResolutionObservationContext";
import { ResolutionSnapshotObserver } from "./resolution/ResolutionSnapshotObserver";
import { ResolveSuggestionResultFactory } from "./resolution/ResolveSuggestionResultFactory";
import { SuggestionLocator } from "./resolution/SuggestionLocator";
import { SuggestionResolutionCleanup } from "./resolution/SuggestionResolutionCleanup";
import { SuggestionResolutionObserver } from "./resolution/SuggestionResolutionObserver";
import { TrackedChangeResolutionExecutor } from "./resolution/TrackedChangeResolutionExecutor";
import {
  getDefaultTextLocator,
  type TextLocator,
} from "./WordTextLocatorContext";

type CohesiveResolutionOutcome = {
  status: SuggestionActionResult["status"];
  trackedChangesAffected: number;
  commentDeleted: boolean;
  pendingBefore: import("../../domain/types").DocumentReviewState;
  pendingAfter: import("../../domain/types").DocumentReviewState;
  executionReport?: ResolutionExecutionReport;
  error?: string;
};

type PreparedExecutionObservation =
  | {
      pendingBefore: import("../../domain/types").DocumentReviewState;
      observation: ResolutionObservation;
    }
  | {
      outcome: CohesiveResolutionOutcome;
    };

type ExecutedResolutionObservation = {
  observation: ResolutionObservation;
  executionReport: ResolutionExecutionReport;
};

// ---------------------------------------------------------------------------
// ResolveSuggestionCommand
// ---------------------------------------------------------------------------

/**
 * Command that resolves (accepts or rejects) a single suggestion in Word.
 *
 * Usage:
 * ```ts
 * const result = await new ResolveSuggestionCommand(suggestion, "accept").execute();
 * ```
 */
export class ResolveSuggestionCommand {
  private readonly stateInspector: DocumentReviewStateInspector;
  private readonly locator: SuggestionLocator;
  private readonly cleanup: SuggestionResolutionCleanup;
  private readonly executor: TrackedChangeResolutionExecutor;
  private readonly replaceResolutionStrategy: ReplaceResolutionStrategy;
  private readonly resultFactory: ResolveSuggestionResultFactory;
  private readonly commentOnlyResolver: CommentOnlySuggestionResolver;
  private readonly observer: SuggestionResolutionObserver;
  private readonly observabilityReporter: ResolutionObservabilityReporter;
  private readonly errorSerializer = new ResolutionErrorSerializer();
  private readonly replaceResolutionWorkflow: ReplaceResolutionWorkflow;
  private readonly snapshotObserver: ResolutionSnapshotObserver;
  private readonly executeStateMachine = new ExecuteResolutionStateMachine();
  private lastExecutionReport?: ResolutionExecutionReport;
  private workflowAttemptId = "";

  constructor(
    private readonly suggestion: Suggestion,
    private readonly action: "accept" | "reject",
    replaceResolutionStrategy: ReplaceResolutionStrategy,
    textLocator: TextLocator = getDefaultTextLocator(),
    observabilityPort: IResolutionObservabilityPort = new NoopResolutionObservabilityAdapter(),
  ) {
    this.stateInspector = new DocumentReviewStateInspector();
    this.locator = new SuggestionLocator(suggestion);
    this.replaceResolutionStrategy = replaceResolutionStrategy;
    this.cleanup = new SuggestionResolutionCleanup(suggestion.id, action);
    this.executor = new TrackedChangeResolutionExecutor(
      suggestion.id,
      action,
      this.replaceResolutionStrategy,
    );
    this.resultFactory = new ResolveSuggestionResultFactory(
      action,
      this.stateInspector,
    );
    this.commentOnlyResolver = new CommentOnlySuggestionResolver(
      suggestion.id,
      this.resultFactory,
      this.stateInspector,
    );
    this.observer = new SuggestionResolutionObserver(
      suggestion,
      this.locator,
      textLocator,
    );
    this.observabilityReporter = new ResolutionObservabilityReporter(
      suggestion.id,
      action,
      observabilityPort,
    );
    this.replaceResolutionWorkflow = new ReplaceResolutionWorkflow(
      action,
      this.locator,
      this.observer,
      this.executor,
      this.replaceResolutionStrategy,
    );
    this.snapshotObserver = new ResolutionSnapshotObserver(
      suggestion,
      this.stateInspector,
      this.locator,
      this.observer,
      this.observabilityReporter,
      this.errorSerializer,
    );
  }

  /**
   * Executes the resolution command.
   * Never throws — catches all errors and returns a result object.
   */
  async execute(): Promise<SuggestionActionResult> {
    try {
      return await Word.run(async (context) => {
        const outcome = await this.executeCohesiveResolution(context);
        return this.resultFactory.buildResolutionResult(
          outcome.status,
          outcome.trackedChangesAffected,
          outcome.commentDeleted,
          outcome.pendingBefore,
          outcome.pendingAfter,
          outcome.error,
          outcome.executionReport,
        );
      });
    } catch (error) {
      if (!this.executeStateMachine.isTerminal) {
        this.executeStateMachine.fail();
      }
      const serializedError = this.errorSerializer.serialize(error);
      const message = serializedError.message;
      console.warn(
        `🧪 [ResolveSuggestionCommand] observe-before failure bracket`,
        {
          workflowAttemptId: this.workflowAttemptId,
          suggestionId: this.suggestion.id,
          action: this.action,
          phase: this.executeStateMachine.currentPhase ?? "unknown",
          execution: this.formatExecutionReportForLog(this.lastExecutionReport),
          error: serializedError,
        },
      );
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" suggestionId="${this.suggestion.id}" action=${this.action} phase=${this.executeStateMachine.currentPhase ?? "unknown"} execution=${this.formatExecutionReportForLog(this.lastExecutionReport)} error=${message}`,
      );
      const pendingAfter = await Word.run((ctx) =>
        this.stateInspector.inspect(ctx),
      ).catch(() => this.stateInspector.buildEmptyState());

      return this.resultFactory.buildErrorResult(
        message,
        pendingAfter,
        this.executeStateMachine.currentPhase ?? undefined,
        this.lastExecutionReport,
      );
    }
  }

  /** Runs the full resolution sequence under a single semantic owner. */
  private async executeCohesiveResolution(
    context: Word.RequestContext,
  ): Promise<CohesiveResolutionOutcome> {
    this.lastExecutionReport = undefined;
    this.workflowAttemptId = this.buildWorkflowAttemptId();
    this.observabilityReporter.setWorkflowAttemptId(this.workflowAttemptId);
    this.transitionExecuteState("locating");
    console.log(
      `🎯 [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" type=${this.suggestion.type}`,
    );
    console.log(
      `🧾 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" suggestion-detail`,
      {
        anchor: this.suggestion.anchor,
        suggestedText: this.suggestion.suggestedText ?? null,
        context: this.suggestion.context,
      },
    );
    const preparation = await this.prepareExecutionObservation(context);
    if ("outcome" in preparation) {
      return preparation.outcome;
    }

    const { pendingBefore, observation } = preparation;
    const executed = await this.executeObservedResolution(context, observation);
    if (executed.executionReport.error) {
      throw new Error(executed.executionReport.error);
    }

    return this.finalizeSuccessfulResolution(
      context,
      pendingBefore,
      executed.observation,
      executed.executionReport,
    );
  }

  /** Locates the suggestion, observes executable evidence, and returns either an early outcome or a ready observation. */
  private async prepareExecutionObservation(
    context: Word.RequestContext,
  ): Promise<PreparedExecutionObservation> {
    await this.observabilityReporter.emitPhase("locate", "started", {
      suggestionType: this.suggestion.type,
    });
    const { rankedCandidates, selectedCc: cc } =
      await this.locator.locateResolutionArtifacts(context);
    console.log(
      `🔎 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" locate candidates=${rankedCandidates.length} selectedCc=${cc?.tag ?? "none"}`,
    );
    await this.observabilityReporter.emitPhase(
      "locate",
      cc ? "succeeded" : "failed",
      {
        candidateCount: rankedCandidates.length,
        selectedCcFound: Boolean(cc),
      },
    );

    if (!cc) {
      this.transitionExecuteState("completed");
      await this.observabilityReporter.emitPhase("observe-before", "failed", {
        reason: "cc-not-found",
      });
      const pendingBefore = await this.stateInspector.inspect(context);
      return {
        outcome: {
          status: "cc-not-found",
          trackedChangesAffected: 0,
          commentDeleted: false,
          pendingBefore,
          pendingAfter: pendingBefore,
        },
      };
    }

    const pendingBefore = await this.stateInspector.inspect(context);
    this.transitionExecuteState("observing-before");
    console.log(`🧪 [ResolveSuggestionCommand] observe-before start`, {
      workflowAttemptId: this.workflowAttemptId,
      suggestionId: this.suggestion.id,
      action: this.action,
      selectedCcTag: cc.tag,
      rankedCandidateCount: rankedCandidates.length,
      suggestionType: this.suggestion.type,
    });
    await this.observabilityReporter.emitPhase("observe-before", "started", {
      suggestionType: this.suggestion.type,
    });

    if (this.suggestion.type === "comment-only") {
      const colocatedComment = await this.locator.findColocatedStylisticComment(
        context,
        cc,
      );
      const commentDeleted = await this.cleanup.deleteLocatedStylisticComment(
        context,
        colocatedComment,
      );
      this.transitionExecuteState("completed");
      await this.observabilityReporter.emitPhase(
        "cleanup-comment",
        "succeeded",
        {
          commentDeleted,
          commentOnly: true,
        },
      );
      const result = await this.commentOnlyResolver.resolve({
        context,
        cc,
        commentDeleted,
        pendingBefore,
      });

      return {
        outcome: {
          status: result.status,
          trackedChangesAffected: result.trackedChangesAffected,
          commentDeleted: result.commentDeleted,
          pendingBefore,
          pendingAfter: result.pendingAfter,
          error: result.error,
          executionReport: result.executionReport,
        },
      };
    }

    if (!this.hasValidTrackChangeContract()) {
      this.transitionExecuteState("completed");
      await this.observabilityReporter.emitPhase("observe-before", "failed", {
        reason: "invalid-track-change-contract",
        suggestionType: this.suggestion.type,
      });
      const invalidContractResult = this.resultFactory.buildErrorResult(
        "Contrato invalido de track-change: anchor y suggestedText son obligatorios.",
        pendingBefore,
        "observe-before",
      );
      return {
        outcome: {
          status: invalidContractResult.status,
          trackedChangesAffected: invalidContractResult.trackedChangesAffected,
          commentDeleted: invalidContractResult.commentDeleted,
          pendingBefore,
          pendingAfter: invalidContractResult.pendingAfter,
          error: invalidContractResult.error,
          executionReport: invalidContractResult.executionReport,
        },
      };
    }

    let observation: ResolutionObservation;
    try {
      observation = await this.observer.observeResolutionCandidates(
        context,
        rankedCandidates,
        cc,
      );
    } catch (error) {
      console.warn(
        `🧪 [ResolveSuggestionCommand] observe-before observer failed`,
        {
          workflowAttemptId: this.workflowAttemptId,
          suggestionId: this.suggestion.id,
          action: this.action,
          selectedCcTag: cc.tag,
          rankedCandidateCount: rankedCandidates.length,
          error: this.errorSerializer.serialize(error),
        },
      );
      throw error;
    }
    console.log(
      `🧭 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" observation status=${observation.observationStatus} trackedChanges=${observation.trackedChanges.length} types=${observation.debugMetadata?.trackedChangeTypes ?? ""} selectedCc=${observation.debugMetadata?.selectedCcTag ?? cc.tag}`,
    );
    console.log(
      `🧾 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" observation-detail`,
      {
        selectedCcTag: observation.selectedCc.tag,
        selectedCommentFound: Boolean(observation.selectedComment),
        trackedChanges: describeTrackedChangesForLog(
          observation.trackedChanges,
        ),
        debugMetadata: observation.debugMetadata ?? null,
      },
    );

    if (observation.observationStatus === "identity-lost") {
      this.transitionExecuteState("completed");
      await this.observabilityReporter.emitPhase(
        "observe-before",
        "failed",
        this.observabilityReporter.mergeMetadata(
          { reason: "identity-lost" },
          observation.debugMetadata,
        ),
      );
      const result = await this.resultFactory.buildObservationFailureResult(
        context,
        "identity-lost",
        pendingBefore,
      );
      return {
        outcome: {
          status: result.status,
          trackedChangesAffected: result.trackedChangesAffected,
          commentDeleted: result.commentDeleted,
          pendingBefore,
          pendingAfter: result.pendingAfter,
          error: result.error,
          executionReport: result.executionReport,
        },
      };
    }

    if (
      observation.observationStatus !== "confirmed-pending" ||
      observation.trackedChanges.length === 0
    ) {
      this.transitionExecuteState("completed");
      await this.observabilityReporter.emitPhase(
        "observe-before",
        "warning",
        this.observabilityReporter.mergeMetadata(
          { reason: "unobservable" },
          observation.debugMetadata,
        ),
      );
      const result = await this.resultFactory.buildObservationFailureResult(
        context,
        "unobservable",
        pendingBefore,
      );
      return {
        outcome: {
          status: result.status,
          trackedChangesAffected: result.trackedChangesAffected,
          commentDeleted: result.commentDeleted,
          pendingBefore,
          pendingAfter: result.pendingAfter,
          error: result.error,
          executionReport: result.executionReport,
        },
      };
    }

    await this.observabilityReporter.emitPhase(
      "observe-before",
      "succeeded",
      this.observabilityReporter.mergeMetadata(
        { trackedChangesObserved: observation.trackedChanges.length },
        observation.debugMetadata,
      ),
    );

    return {
      pendingBefore,
      observation,
    };
  }

  /** Cleans the resolved artifacts, inspects final state, and assembles the success outcome. */
  private async finalizeSuccessfulResolution(
    context: Word.RequestContext,
    pendingBefore: import("../../domain/types").DocumentReviewState,
    observation: ResolutionObservation,
    executionReport: ResolutionExecutionReport,
  ): Promise<CohesiveResolutionOutcome> {
    this.transitionExecuteState("cleaning-comment");
    console.log(
      `🧹 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" cleanup-comment selected=${Boolean(observation.selectedComment)}`,
    );
    const commentDeleted =
      await this.cleanup.deleteLocatedStylisticCommentAfterResolution(
        context,
        observation.selectedComment,
      );
    await this.observabilityReporter.emitPhase("cleanup-comment", "succeeded", {
      commentDeleted,
    });

    this.transitionExecuteState("cleaning-anchor");
    console.log(
      `🧹 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" cleanup-anchor cc="${observation.selectedCc.tag}"`,
    );
    await this.cleanup.cleanupResolvedSuggestionAnchor(
      context,
      observation.selectedCc,
    );
    await this.observabilityReporter.emitPhase("cleanup-anchor", "succeeded", {
      anchorDeleted: true,
    });

    this.transitionExecuteState("inspecting-after");
    const pendingAfter =
      await this.stateInspector.inspectAfterResolution(context);
    await this.snapshotObserver.capture(
      context,
      "after-cleanup-before-return",
      observation.selectedCc,
      pendingAfter,
    );
    await this.observabilityReporter.emitPhase("inspect-after", "succeeded", {
      pendingArtifacts: pendingAfter.pendingStylisticArtifacts,
    });
    this.transitionExecuteState("completed");

    return {
      status: this.resultFactory.toResolutionStatus(),
      trackedChangesAffected: executionReport.completed,
      commentDeleted,
      pendingBefore,
      pendingAfter,
      executionReport,
    };
  }

  /** Executes the observed tracked changes, folds in post-execute observation, and captures the final execution report. */
  private async executeObservedResolution(
    context: Word.RequestContext,
    observation: ResolutionObservation,
  ): Promise<ExecutedResolutionObservation> {
    this.transitionExecuteState("executing");
    await this.observabilityReporter.emitPhase("execute", "started", {
      trackedChangesAttempted: observation.trackedChanges.length,
      trackedChangeTypes: observation.debugMetadata?.trackedChangeTypes ?? "",
    });

    const executeAttempt = await this.replaceResolutionWorkflow.execute(
      context,
      observation,
      this.workflowAttemptId,
    );
    const postExecuteAttempt = await this.observePostExecuteResolutionState(
      context,
      executeAttempt.observation.selectedCc,
    );
    const executionReport = this.mergePostExecutePendingError(
      executeAttempt.executionReport,
      postExecuteAttempt.observation,
    );
    this.lastExecutionReport = executionReport;
    console.log(
      `🧾 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" execute-result-detail`,
      {
        executionReport,
        observation: {
          selectedCcTag: executeAttempt.observation.selectedCc.tag,
          trackedChanges: describeTrackedChangesForLog(
            executeAttempt.observation.trackedChanges,
          ),
          debugMetadata: executeAttempt.observation.debugMetadata ?? null,
        },
      },
    );
    await this.observabilityReporter.emitPhase(
      "execute",
      executionReport.error ? "failed" : "succeeded",
      {
        attempted: executionReport.attempted,
        completed: executionReport.completed,
        remaining: executionReport.remaining,
        failureIndex: executionReport.failureIndex ?? null,
        recoveryAttempted: executeAttempt.recoveryAttempted,
        recoverySucceeded: executeAttempt.recoverySucceeded,
        postExecuteTrackedChangesObserved:
          postExecuteAttempt.observation?.trackedChanges.length ?? null,
        postExecuteObservationStatus:
          postExecuteAttempt.observation?.observationStatus ?? null,
        postExecuteRecoveryAttempted: postExecuteAttempt.recoveryAttempted,
        postExecuteRecoverySucceeded: postExecuteAttempt.recoverySucceeded,
      },
    );

    return {
      observation: executeAttempt.observation,
      executionReport,
    };
  }

  /** Returns true when the current tracked-change suggestion satisfies the contract. */
  private hasValidTrackChangeContract(): boolean {
    return (
      this.suggestion.type === "track-change" &&
      this.suggestion.anchor.trim().length > 0 &&
      (this.suggestion.suggestedText?.trim().length ?? 0) > 0
    );
  }

  /** Re-checks post-execute replace state without executing any fallback recovery. */
  private async observePostExecuteResolutionState(
    context: Word.RequestContext,
    preferredCc: Word.ContentControl,
  ): Promise<{
    observation: ResolutionObservation | null;
    recoveryAttempted: boolean;
    recoverySucceeded: boolean;
  }> {
    const postExecuteObservation = await this.snapshotObserver.capture(
      context,
      "after-execute-before-cleanup",
      preferredCc,
    );

    if (!postExecuteObservation) {
      return {
        observation: null,
        recoveryAttempted: false,
        recoverySucceeded: false,
      };
    }

    return {
      observation: postExecuteObservation,
      recoveryAttempted: false,
      recoverySucceeded: false,
    };
  }

  /** Fails closed when a fresh full observation still sees the replace pending after execute. */
  private mergePostExecutePendingError(
    executionReport: ResolutionExecutionReport,
    postExecuteObservation: ResolutionObservation | null,
  ): ResolutionExecutionReport {
    if (
      !postExecuteObservation ||
      postExecuteObservation.trackedChanges.length === 0
    ) {
      return executionReport;
    }

    // A post-execute snapshot may include adjacent-neighbor TCs via
    // `bodyRelated` (AdjacentBefore/AdjacentAfter). Those belong to other
    // suggestions and must NOT be reported as this suggestion's pending
    // remainder. Only CC-scoped TCs (inside cc or ccRange) count.
    const ccScopedRemaining =
      (postExecuteObservation.debugMetadata?.ccTrackedChangesCount ?? 0) +
      (postExecuteObservation.debugMetadata?.ccRangeTrackedChangesCount ?? 0);
    if (ccScopedRemaining === 0) {
      return executionReport;
    }

    return {
      ...executionReport,
      remaining: ccScopedRemaining,
      error:
        "Word siguió exponiendo tracked changes del replace después de ejecutar la resolución; se cancela cleanup para evitar falso success.",
    };
  }

  /** Builds a correlation id for one resolution workflow attempt. */
  private buildWorkflowAttemptId(): string {
    return `${this.suggestion.id}:${this.action}:${Date.now()}`;
  }

  /** Advances the internal execute state machine. */
  private transitionExecuteState(
    next: import("../../domain/suggestion/ExecuteResolutionStateMachine").ExecuteResolutionState,
  ): void {
    this.executeStateMachine.transition(next);
  }

  /** Formats one execution report for compact console diagnostics. */
  private formatExecutionReportForLog(
    report?: ResolutionExecutionReport,
  ): string {
    if (!report) {
      return "none";
    }

    return `attempted=${report.attempted},completed=${report.completed},remaining=${report.remaining},failureIndex=${report.failureIndex ?? "null"},error=${report.error ?? "none"}`;
  }
}
