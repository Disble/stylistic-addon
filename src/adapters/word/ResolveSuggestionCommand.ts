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

import type { ITelemetryPort } from "../../domain/ports";
import { ExecuteResolutionStateMachine } from "../../domain/suggestion/ExecuteResolutionStateMachine";
import type {
  ResolutionExecutionReport,
  ResolutionPhase,
  ResolutionTelemetryEvent,
  Suggestion,
  SuggestionActionResult,
} from "../../domain/types";
import { CommentOnlySuggestionResolver } from "./resolution/CommentOnlySuggestionResolver";
import { DocumentReviewStateInspector } from "./resolution/DocumentReviewStateInspector";
import type { ResolutionObservation } from "./resolution/ResolutionContext";
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

type ResolutionTelemetryMetadata = Record<
  string,
  string | number | boolean | null
>;

type TrackedChangeLogEntry = {
  id: string;
  type: string;
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
  private readonly resultFactory: ResolveSuggestionResultFactory;
  private readonly commentOnlyResolver: CommentOnlySuggestionResolver;
  private readonly observer: SuggestionResolutionObserver;
  private readonly executeStateMachine = new ExecuteResolutionStateMachine();
  private lastExecutionReport?: ResolutionExecutionReport;
  private workflowAttemptId = "";

  constructor(
    private readonly suggestion: Suggestion,
    private readonly action: "accept" | "reject",
    textLocator: TextLocator = getDefaultTextLocator(),
    private readonly telemetryPort: ITelemetryPort = {
      emit: async () => undefined,
    },
  ) {
    this.stateInspector = new DocumentReviewStateInspector();
    this.locator = new SuggestionLocator(suggestion);
    this.cleanup = new SuggestionResolutionCleanup(suggestion.id, action);
    this.executor = new TrackedChangeResolutionExecutor(suggestion.id, action);
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
      const message = error instanceof Error ? error.message : String(error);
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
    await this.emitTelemetry("locate", "started", {
      suggestionType: this.suggestion.type,
    });
    const { rankedCandidates, selectedCc: cc } =
      await this.locator.locateResolutionArtifacts(context);
    console.log(
      `🔎 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" locate candidates=${rankedCandidates.length} selectedCc=${cc?.tag ?? "none"}`,
    );
    await this.emitTelemetry("locate", cc ? "succeeded" : "failed", {
      candidateCount: rankedCandidates.length,
      selectedCcFound: Boolean(cc),
    });

    if (!cc) {
      this.transitionExecuteState("completed");
      await this.emitTelemetry("observe-before", "failed", {
        reason: "cc-not-found",
      });
      const pendingBefore = await this.stateInspector.inspect(context);
      return {
        status: "cc-not-found",
        trackedChangesAffected: 0,
        commentDeleted: false,
        pendingBefore,
        pendingAfter: pendingBefore,
      };
    }

    const pendingBefore = await this.stateInspector.inspect(context);
    this.transitionExecuteState("observing-before");
    await this.emitTelemetry("observe-before", "started", {
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
      await this.emitTelemetry("cleanup-comment", "succeeded", {
        commentDeleted,
        commentOnly: true,
      });
      const result = await this.commentOnlyResolver.resolve({
        context,
        cc,
        commentDeleted,
        pendingBefore,
      });

      return {
        status: result.status,
        trackedChangesAffected: result.trackedChangesAffected,
        commentDeleted: result.commentDeleted,
        pendingBefore,
        pendingAfter: result.pendingAfter,
        error: result.error,
        executionReport: result.executionReport,
      };
    }

    const observation = await this.observer.observeResolutionCandidates(
      context,
      rankedCandidates,
      cc,
    );
    console.log(
      `🧭 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" observation status=${observation.observationStatus} trackedChanges=${observation.trackedChanges.length} types=${observation.debugMetadata?.trackedChangeTypes ?? ""} selectedCc=${observation.debugMetadata?.selectedCcTag ?? cc.tag}`,
    );
    console.log(
      `🧾 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" observation-detail`,
      {
        selectedCcTag: observation.selectedCc.tag,
        selectedCommentFound: Boolean(observation.selectedComment),
        trackedChanges: this.describeTrackedChangesForLog(
          observation.trackedChanges,
        ),
        debugMetadata: observation.debugMetadata ?? null,
      },
    );

    if (observation.observationStatus === "identity-lost") {
      this.transitionExecuteState("completed");
      await this.emitTelemetry(
        "observe-before",
        "failed",
        this.mergeTelemetryMetadata(
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
        status: result.status,
        trackedChangesAffected: result.trackedChangesAffected,
        commentDeleted: result.commentDeleted,
        pendingBefore,
        pendingAfter: result.pendingAfter,
        error: result.error,
        executionReport: result.executionReport,
      };
    }

    if (
      observation.observationStatus !== "confirmed-pending" ||
      observation.trackedChanges.length === 0
    ) {
      this.transitionExecuteState("completed");
      await this.emitTelemetry(
        "observe-before",
        "warning",
        this.mergeTelemetryMetadata(
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
        status: result.status,
        trackedChangesAffected: result.trackedChangesAffected,
        commentDeleted: result.commentDeleted,
        pendingBefore,
        pendingAfter: result.pendingAfter,
        error: result.error,
        executionReport: result.executionReport,
      };
    }

    await this.emitTelemetry(
      "observe-before",
      "succeeded",
      this.mergeTelemetryMetadata(
        { trackedChangesObserved: observation.trackedChanges.length },
        observation.debugMetadata,
      ),
    );
    this.transitionExecuteState("executing");
    await this.emitTelemetry("execute", "started", {
      trackedChangesAttempted: observation.trackedChanges.length,
      trackedChangeTypes: observation.debugMetadata?.trackedChangeTypes ?? "",
    });

    const executeAttempt =
      await this.executeTrackedChangesWithImmediateRecovery(
        context,
        observation,
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
          trackedChanges: this.describeTrackedChangesForLog(
            executeAttempt.observation.trackedChanges,
          ),
          debugMetadata: executeAttempt.observation.debugMetadata ?? null,
        },
      },
    );
    await this.emitTelemetry(
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

    if (executionReport.error) {
      throw new Error(executionReport.error);
    }

    this.transitionExecuteState("cleaning-comment");
    console.log(
      `🧹 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" cleanup-comment selected=${Boolean(executeAttempt.observation.selectedComment)}`,
    );
    const commentDeleted =
      await this.cleanup.deleteLocatedStylisticCommentAfterResolution(
        context,
        executeAttempt.observation.selectedComment,
      );
    await this.emitTelemetry("cleanup-comment", "succeeded", {
      commentDeleted,
    });

    this.transitionExecuteState("cleaning-anchor");
    console.log(
      `🧹 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" cleanup-anchor cc="${executeAttempt.observation.selectedCc.tag}"`,
    );
    await this.cleanup.cleanupResolvedSuggestionAnchor(
      context,
      executeAttempt.observation.selectedCc,
    );
    await this.emitTelemetry("cleanup-anchor", "succeeded", {
      anchorDeleted: true,
    });

    this.transitionExecuteState("inspecting-after");
    const pendingAfter =
      await this.stateInspector.inspectAfterResolution(context);
    await this.logWorkflowSnapshot(
      context,
      "after-cleanup-before-return",
      executeAttempt.observation.selectedCc,
      pendingAfter,
    );
    await this.emitTelemetry("inspect-after", "succeeded", {
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

  /** Executes tracked changes and retries once after partial progress if re-observation can recover the remainder. */
  private async executeTrackedChangesWithImmediateRecovery(
    context: Word.RequestContext,
    observation: ResolutionObservation,
  ): Promise<{
    observation: ResolutionObservation;
    executionReport: ResolutionExecutionReport;
    recoveryAttempted: boolean;
    recoverySucceeded: boolean;
  }> {
    if (this.isReplaceSuggestion()) {
      return this.executeReplaceTrackedChangesWithReobservation(
        context,
        observation,
      );
    }

    console.log(
      `⚙️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" execute-start trackedChanges=${observation.trackedChanges.length} types=${observation.debugMetadata?.trackedChangeTypes ?? ""}`,
    );
    const initialReport = await this.executor.apply(
      context,
      observation.trackedChanges,
    );

    if (!initialReport.error || initialReport.remaining === 0) {
      return {
        observation,
        executionReport: initialReport,
        recoveryAttempted: false,
        recoverySucceeded: false,
      };
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" attempting same-click recovery after partial execute failure: ${initialReport.error}`,
    );

    try {
      const relocated = await this.locator.locateResolutionArtifacts(context);
      if (!relocated.selectedCc) {
        return {
          observation,
          executionReport: initialReport,
          recoveryAttempted: true,
          recoverySucceeded: false,
        };
      }

      const recoveryObservation =
        await this.observer.observeResolutionCandidates(
          context,
          relocated.rankedCandidates,
          relocated.selectedCc,
        );
      console.log(
        `🔁 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" recovery observation status=${recoveryObservation.observationStatus} trackedChanges=${recoveryObservation.trackedChanges.length} types=${recoveryObservation.debugMetadata?.trackedChangeTypes ?? ""}`,
      );

      const canContinueRecovery =
        recoveryObservation.observationStatus !== "identity-lost" &&
        recoveryObservation.trackedChanges.length > 0;

      if (!canContinueRecovery) {
        return {
          observation,
          executionReport: initialReport,
          recoveryAttempted: true,
          recoverySucceeded: false,
        };
      }

      const recoveryReport = await this.executor.apply(
        context,
        recoveryObservation.trackedChanges,
      );

      return {
        observation: recoveryObservation,
        executionReport: this.mergeExecutionReports(
          initialReport,
          recoveryReport,
        ),
        recoveryAttempted: true,
        recoverySucceeded: !recoveryReport.error,
      };
    } catch (recoveryError) {
      const message =
        recoveryError instanceof Error
          ? recoveryError.message
          : String(recoveryError);

      console.warn(
        `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" same-click recovery failed: ${message}`,
      );

      return {
        observation,
        executionReport: initialReport,
        recoveryAttempted: true,
        recoverySucceeded: false,
      };
    }
  }

  /** Resolves replace suggestions in two semantic host steps with fresh re-observation between them. */
  private async executeReplaceTrackedChangesWithReobservation(
    context: Word.RequestContext,
    observation: ResolutionObservation,
  ): Promise<{
    observation: ResolutionObservation;
    executionReport: ResolutionExecutionReport;
    recoveryAttempted: boolean;
    recoverySucceeded: boolean;
  }> {
    if (this.hasDuplicateReplaceSide(observation.trackedChanges)) {
      const normalizedObservation =
        this.normalizeReplaceObservation(observation);
      console.log(
        `⚙️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" duplicate-side replace normalized to semantic pair types=${normalizedObservation.debugMetadata?.trackedChangeTypes ?? ""}`,
      );
      return this.executeReplaceTrackedChangesWithReobservation(
        context,
        normalizedObservation,
      );
    }

    const semanticOrder = this.getReplaceSemanticOrder();
    let activeObservation = observation;
    let completed = 0;
    let recoveryAttempted = false;
    let recoverySucceeded = false;

    for (const [stepIndex, trackedChangeType] of semanticOrder.entries()) {
      const stepResult = await this.executeReplaceSemanticStep(
        context,
        activeObservation,
        trackedChangeType,
      );

      recoveryAttempted ||= stepResult.recoveryAttempted;
      recoverySucceeded ||= stepResult.recoverySucceeded;
      activeObservation = stepResult.observation;

      if (stepResult.completed) {
        completed += 1;
      }

      if (stepResult.error) {
        return this.buildReplaceFailureOutcome(
          await this.tryAtomicAcceptReplaceFallback(
            context,
            stepIndex,
            activeObservation,
            stepResult.observation,
            completed,
          ),
          {
            observation: activeObservation,
            attempted: semanticOrder.length,
            completed,
            failureIndex: stepIndex,
            error: stepResult.error,
            recoveryAttempted,
            recoverySucceeded,
          },
        );
      }

      if (stepIndex < semanticOrder.length - 1) {
        const reobservationResult = await this.reobserveRemainingReplaceSide(
          context,
          activeObservation,
          semanticOrder,
          stepIndex,
          completed,
          recoveryAttempted,
          recoverySucceeded,
        );
        if ("executionReport" in reobservationResult) {
          return reobservationResult;
        }

        activeObservation = reobservationResult.observation;
      }
    }

    return {
      observation: activeObservation,
      executionReport: {
        attempted: semanticOrder.length,
        completed,
        remaining: semanticOrder.length - completed,
      },
      recoveryAttempted,
      recoverySucceeded,
    };
  }

  /** Falls back to the pre-refactor atomic accept behavior when Deleted-alone execution is rejected by Word. */
  private async tryAtomicAcceptReplaceFallback(
    context: Word.RequestContext,
    stepIndex: number,
    initialObservation: ResolutionObservation,
    latestObservation: ResolutionObservation,
    completed: number,
  ): Promise<{
    observation: ResolutionObservation;
    executionReport: ResolutionExecutionReport;
  } | null> {
    if (
      this.action !== "accept" ||
      stepIndex !== 0 ||
      completed !== 0 ||
      initialObservation.trackedChanges.length < 2 ||
      latestObservation.trackedChanges.length < 2
    ) {
      return null;
    }

    const hasDeleted =
      this.findTrackedChangeByType(
        latestObservation.trackedChanges,
        "Deleted",
      ) !== null;
    const hasAdded =
      this.findTrackedChangeByType(
        latestObservation.trackedChanges,
        "Added",
      ) !== null;

    if (!hasDeleted || !hasAdded) {
      return null;
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" retrying accept replace as one atomic batch after Deleted-alone execution failed`,
    );

    const atomicReport = await this.executor.applyAtomically(
      context,
      latestObservation.trackedChanges,
    );
    if (atomicReport.error) {
      return null;
    }

    return {
      observation: latestObservation,
      executionReport: atomicReport,
    };
  }

  /** Tries the known-good atomic accept behavior once a fresh observation still exposes the full replace pair. */
  private async tryAtomicAcceptStepFallback(
    context: Word.RequestContext,
    trackedChangeType: "Added" | "Deleted",
    observation: ResolutionObservation,
  ): Promise<{
    observation: ResolutionObservation;
    completed: boolean;
    error?: string;
    recoveryAttempted: boolean;
    recoverySucceeded: boolean;
  } | null> {
    if (
      this.action !== "accept" ||
      trackedChangeType !== "Deleted" ||
      observation.trackedChanges.length < 2
    ) {
      return null;
    }

    const hasDeleted =
      this.findTrackedChangeByType(observation.trackedChanges, "Deleted") !==
      null;
    const hasAdded =
      this.findTrackedChangeByType(observation.trackedChanges, "Added") !==
      null;

    if (!hasDeleted || !hasAdded) {
      return null;
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" retrying accept replace as one atomic batch after Deleted-alone execution failed`,
    );

    const atomicReport = await this.executor.applyAtomically(
      context,
      observation.trackedChanges,
    );
    if (atomicReport.error) {
      return null;
    }

    return {
      observation,
      completed: true,
      recoveryAttempted: true,
      recoverySucceeded: true,
    };
  }

  /** Executes one semantic side of a replace and retries with fresh proxies if needed. */
  private async executeReplaceSemanticStep(
    context: Word.RequestContext,
    observation: ResolutionObservation,
    trackedChangeType: "Added" | "Deleted",
  ): Promise<{
    observation: ResolutionObservation;
    completed: boolean;
    error?: string;
    recoveryAttempted: boolean;
    recoverySucceeded: boolean;
  }> {
    const trackedChange = this.findTrackedChangeByType(
      observation.trackedChanges,
      trackedChangeType,
    );

    if (!trackedChange) {
      return {
        observation,
        completed: false,
        error: `Word no reexpuso el tracked change ${trackedChangeType} requerido para resolver el replace.`,
        recoveryAttempted: false,
        recoverySucceeded: false,
      };
    }

    const initialReport = await this.executor.apply(context, [trackedChange]);
    if (!initialReport.error) {
      return {
        observation,
        completed: true,
        recoveryAttempted: false,
        recoverySucceeded: false,
      };
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace-step=${trackedChangeType} retrying after failure: ${initialReport.error}`,
    );

    const firstRecoveryObservation = await this.reobserveResolutionCandidates(
      context,
      observation.selectedCc,
    );
    if (!firstRecoveryObservation) {
      return {
        observation,
        completed: false,
        error: initialReport.error,
        recoveryAttempted: true,
        recoverySucceeded: false,
      };
    }

    console.log(
      `🔁 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace-step=${trackedChangeType} recovery observation status=${firstRecoveryObservation.observation.observationStatus} trackedChanges=${firstRecoveryObservation.observation.trackedChanges.length} types=${firstRecoveryObservation.observation.debugMetadata?.trackedChangeTypes ?? ""}`,
    );

    const recoveredTrackedChange = this.findTrackedChangeByType(
      firstRecoveryObservation.observation.trackedChanges,
      trackedChangeType,
    );
    const atomicFallback = await this.tryAtomicAcceptStepFallback(
      context,
      trackedChangeType,
      firstRecoveryObservation.observation,
    );
    if (atomicFallback) {
      return atomicFallback;
    }

    if (!recoveredTrackedChange) {
      if (firstRecoveryObservation.observation.trackedChanges.length === 0) {
        return {
          observation: firstRecoveryObservation.observation,
          completed: true,
          recoveryAttempted: true,
          recoverySucceeded: true,
        };
      }

      return {
        observation: firstRecoveryObservation.observation,
        completed: false,
        error:
          "Word dejó el replace en un estado parcial; se cancela cleanup para evitar falso success.",
        recoveryAttempted: true,
        recoverySucceeded: false,
      };
    }

    const recoveryReport = await this.executor.apply(context, [
      recoveredTrackedChange,
    ]);
    if (!recoveryReport.error) {
      return {
        observation: firstRecoveryObservation.observation,
        completed: true,
        recoveryAttempted: true,
        recoverySucceeded: true,
      };
    }

    const finalRecoveryObservation = await this.reobserveResolutionCandidates(
      context,
      firstRecoveryObservation.observation.selectedCc,
    );
    if (!finalRecoveryObservation) {
      return {
        observation: firstRecoveryObservation.observation,
        completed: false,
        error: recoveryReport.error,
        recoveryAttempted: true,
        recoverySucceeded: false,
      };
    }

    console.log(
      `🔁 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace-step=${trackedChangeType} final recovery observation status=${finalRecoveryObservation.observation.observationStatus} trackedChanges=${finalRecoveryObservation.observation.trackedChanges.length} types=${finalRecoveryObservation.observation.debugMetadata?.trackedChangeTypes ?? ""}`,
    );

    const stillPendingTrackedChange = this.findTrackedChangeByType(
      finalRecoveryObservation.observation.trackedChanges,
      trackedChangeType,
    );

    return {
      observation: finalRecoveryObservation.observation,
      completed:
        stillPendingTrackedChange === null &&
        finalRecoveryObservation.observation.trackedChanges.length === 0,
      ...this.buildReplaceRecoveryError(
        stillPendingTrackedChange,
        finalRecoveryObservation.observation.trackedChanges.length,
        recoveryReport.error,
      ),
      recoveryAttempted: true,
      recoverySucceeded:
        stillPendingTrackedChange === null &&
        finalRecoveryObservation.observation.trackedChanges.length === 0,
    };
  }

  /** Re-locates the current suggestion and re-observes it from fresh Word proxies. */
  private async reobserveResolutionCandidates(
    context: Word.RequestContext,
    preferredCc?: Word.ContentControl,
  ): Promise<{
    rankedCandidates: Word.ContentControl[];
    observation: ResolutionObservation;
  } | null> {
    const relocated = await this.locator.locateResolutionArtifacts(context);
    if (!relocated.selectedCc) {
      return null;
    }

    const resolvedPreferredCc = this.resolveFreshPreferredCandidate(
      relocated.rankedCandidates,
      preferredCc,
    );
    const preferredCandidates = this.prioritizeFreshPreferredCandidate(
      relocated.rankedCandidates,
      resolvedPreferredCc,
    );

    const reobserved = await this.observer.observeResolutionCandidates(
      context,
      preferredCandidates,
      resolvedPreferredCc ?? relocated.selectedCc,
    );

    return {
      rankedCandidates: relocated.rankedCandidates,
      observation: reobserved,
    };
  }

  /** Re-locates the current suggestion and re-observes only one semantic replace side. */
  private async reobserveResolutionCandidatesForSemanticSide(
    context: Word.RequestContext,
    trackedChangeType: "Added" | "Deleted",
    preferredCc?: Word.ContentControl,
  ): Promise<{
    rankedCandidates: Word.ContentControl[];
    observation: ResolutionObservation;
  } | null> {
    const relocated = await this.locator.locateResolutionArtifacts(context);
    if (!relocated.selectedCc) {
      return null;
    }

    const resolvedPreferredCc = this.resolveFreshPreferredCandidate(
      relocated.rankedCandidates,
      preferredCc,
    );
    const preferredCandidates = this.prioritizeFreshPreferredCandidate(
      relocated.rankedCandidates,
      resolvedPreferredCc,
    );

    const reobserved = this.isReplaceSuggestion()
      ? await this.observer.observeResolutionCandidatesForSemanticSide(
          context,
          preferredCandidates,
          resolvedPreferredCc ?? relocated.selectedCc,
          trackedChangeType,
        )
      : await this.observer.observeResolutionCandidates(
          context,
          preferredCandidates,
          resolvedPreferredCc ?? relocated.selectedCc,
        );

    return {
      rankedCandidates: relocated.rankedCandidates,
      observation: reobserved,
    };
  }

  /** Resolves one stale preferred CC to its fresh logical equivalent from the current locate pass. */
  private resolveFreshPreferredCandidate(
    rankedCandidates: Word.ContentControl[],
    preferredCc?: Word.ContentControl,
  ): Word.ContentControl | null {
    if (!preferredCc) {
      return null;
    }

    const preferredTag = preferredCc.tag;
    const preferredTitle = preferredCc.title ?? "";

    return (
      rankedCandidates.find(
        (candidate) =>
          candidate.tag === preferredTag &&
          (candidate.title ?? "") === preferredTitle,
      ) ?? null
    );
  }

  /** Keeps the fresh logical successor first without ever reusing the old proxy object. */
  private prioritizeFreshPreferredCandidate(
    rankedCandidates: Word.ContentControl[],
    preferredCc: Word.ContentControl | null,
  ): Word.ContentControl[] {
    if (!preferredCc) {
      return rankedCandidates;
    }

    return [
      preferredCc,
      ...rankedCandidates.filter((candidate) => candidate !== preferredCc),
    ];
  }

  /** Returns the first tracked change for the requested semantic side. */
  private findTrackedChangeByType(
    trackedChanges: Word.TrackedChange[],
    trackedChangeType: "Added" | "Deleted",
  ): Word.TrackedChange | null {
    return (
      trackedChanges.find(
        (trackedChange) => trackedChange.type === trackedChangeType,
      ) ?? null
    );
  }

  /** Collapses duplicate replace proxies into one semantic Deleted/Added pair before execution. */
  private normalizeReplaceObservation(
    observation: ResolutionObservation,
  ): ResolutionObservation {
    const normalizedTrackedChanges = [
      this.findTrackedChangeByType(observation.trackedChanges, "Deleted"),
      this.findTrackedChangeByType(observation.trackedChanges, "Added"),
    ].filter(
      (trackedChange): trackedChange is Word.TrackedChange =>
        trackedChange !== null,
    );

    return {
      ...observation,
      trackedChanges: normalizedTrackedChanges,
      debugMetadata: observation.debugMetadata
        ? {
            ...observation.debugMetadata,
            trackedChangesObserved: normalizedTrackedChanges.length,
            trackedChangeTypes: normalizedTrackedChanges
              .map((trackedChange) => trackedChange.type ?? "unknown")
              .join(","),
          }
        : observation.debugMetadata,
    };
  }

  /** Returns true when Word exposes more than one tracked change for the same replace side. */
  private hasDuplicateReplaceSide(
    trackedChanges: Word.TrackedChange[],
  ): boolean {
    let deletedCount = 0;
    let addedCount = 0;

    for (const trackedChange of trackedChanges) {
      if (trackedChange.type === "Deleted") {
        deletedCount += 1;
      }

      if (trackedChange.type === "Added") {
        addedCount += 1;
      }

      if (deletedCount > 1 || addedCount > 1) {
        return true;
      }
    }

    return false;
  }

  /** Queues the observed tracked changes in Word order, matching the pre-refactor duplicate-side path. */
  private queueObservedTrackedChangeResolution(
    trackedChanges: Word.TrackedChange[],
  ): void {
    console.log(
      `🧾 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" queued-observed-tracked-changes`,
      {
        trackedChanges: this.describeTrackedChangesForLog(trackedChanges),
      },
    );
    for (const trackedChange of trackedChanges) {
      if (this.action === "accept") {
        trackedChange.accept();
      } else {
        trackedChange.reject();
      }
    }

    console.log(
      `⚙️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" queued ${this.action} for ${trackedChanges.length} observed tracked changes without immediate replace sync`,
    );
  }

  /** Returns true when the current suggestion is a tracked replace. */
  private isReplaceSuggestion(): boolean {
    return (
      this.suggestion.type === "track-change" &&
      this.suggestion.anchor.length > 0 &&
      (this.suggestion.suggestedText?.length ?? 0) > 0
    );
  }

  /** Merges the initial partial execution report with one immediate recovery pass. */
  private mergeExecutionReports(
    initialReport: ResolutionExecutionReport,
    recoveryReport: ResolutionExecutionReport,
  ): ResolutionExecutionReport {
    const completed = initialReport.completed + recoveryReport.completed;
    const remaining = recoveryReport.remaining;
    const failureIndex = recoveryReport.failureIndex;

    return {
      attempted: completed + remaining,
      completed,
      remaining,
      ...(failureIndex === undefined
        ? {}
        : {
            failureIndex: initialReport.completed + failureIndex,
          }),
      ...(recoveryReport.error ? { error: recoveryReport.error } : {}),
    };
  }

  /** Merges telemetry metadata without using empty-object spread fallbacks inline. */
  private mergeTelemetryMetadata(
    base?: ResolutionTelemetryMetadata,
    extra?: ResolutionTelemetryMetadata,
  ): ResolutionTelemetryMetadata {
    if (!base) {
      return extra ?? {};
    }

    if (!extra) {
      return base;
    }

    return {
      ...base,
      ...extra,
    };
  }

  /**
   * Returns the semantic execution order for replace suggestions.
   *
   * For BOTH accept and reject we resolve the Deleted side first, then the
   * Added side. Rationale: the suggestion CC wraps the Added (inserted) text.
   * Rejecting the Added TC first in Word removes the inserted text and with
   * it the CC, so the inter-step re-observation for the remaining Deleted
   * side fails with `getByTag returned 0 CC candidate(s)` and the workflow
   * reports a false "no pudo reubicar" error. Resolving the Deleted side
   * first keeps the CC intact through the first step for both actions.
   */
  private getReplaceSemanticOrder():
    | readonly ["Deleted", "Added"]
    | readonly ["Added", "Deleted"] {
    return ["Deleted", "Added"] as const;
  }

  /** Builds the terminal outcome for a replace-step failure, preserving one-shot atomic fallback. */
  private buildReplaceFailureOutcome(
    atomicFallback: {
      observation: ResolutionObservation;
      executionReport: ResolutionExecutionReport;
    } | null,
    options: {
      observation: ResolutionObservation;
      attempted: number;
      completed: number;
      failureIndex: number;
      error: string;
      recoveryAttempted: boolean;
      recoverySucceeded: boolean;
    },
  ): {
    observation: ResolutionObservation;
    executionReport: ResolutionExecutionReport;
    recoveryAttempted: boolean;
    recoverySucceeded: boolean;
  } {
    if (atomicFallback) {
      return {
        observation: atomicFallback.observation,
        executionReport: atomicFallback.executionReport,
        recoveryAttempted: true,
        recoverySucceeded: true,
      };
    }

    return {
      observation: options.observation,
      executionReport: {
        attempted: options.attempted,
        completed: options.completed,
        remaining: options.attempted - options.completed,
        failureIndex: options.failureIndex,
        error: options.error,
      },
      recoveryAttempted: options.recoveryAttempted,
      recoverySucceeded: options.recoverySucceeded,
    };
  }

  /** Re-observes only the remaining replace side and rejects any reappearance of the resolved side. */
  private async reobserveRemainingReplaceSide(
    context: Word.RequestContext,
    activeObservation: ResolutionObservation,
    semanticOrder:
      | readonly ["Deleted", "Added"]
      | readonly ["Added", "Deleted"],
    stepIndex: number,
    completed: number,
    recoveryAttempted: boolean,
    recoverySucceeded: boolean,
  ): Promise<
    | {
        observation: ResolutionObservation;
      }
    | {
        observation: ResolutionObservation;
        executionReport: ResolutionExecutionReport;
        recoveryAttempted: boolean;
        recoverySucceeded: boolean;
      }
  > {
    const remainingTrackedChangeType = semanticOrder[stepIndex + 1];
    const resolvedTrackedChangeType = semanticOrder[stepIndex];
    const reobserved = await this.reobserveResolutionCandidatesForSemanticSide(
      context,
      remainingTrackedChangeType,
      activeObservation.selectedCc,
    );

    if (!reobserved) {
      return {
        observation: activeObservation,
        executionReport: {
          attempted: semanticOrder.length,
          completed,
          remaining: semanticOrder.length - completed,
          failureIndex: stepIndex + 1,
          error:
            "Word no pudo reubicar la sugerencia después del primer paso del replace.",
        },
        recoveryAttempted,
        recoverySucceeded,
      };
    }

    console.log(
      `🔁 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace re-observation step=${stepIndex + 1} side=${remainingTrackedChangeType} status=${reobserved.observation.observationStatus} trackedChanges=${reobserved.observation.trackedChanges.length} types=${reobserved.observation.debugMetadata?.trackedChangeTypes ?? ""} semanticSource=${reobserved.observation.debugMetadata?.selectedSemanticSideSource ?? ""} deletedSource=${reobserved.observation.debugMetadata?.selectedDeletedSource ?? ""} addedSource=${reobserved.observation.debugMetadata?.selectedAddedSource ?? ""}`,
    );

    const resolvedSideStillPending = this.findTrackedChangeByType(
      reobserved.observation.trackedChanges,
      resolvedTrackedChangeType,
    );
    if (resolvedSideStillPending) {
      return {
        observation: reobserved.observation,
        executionReport: {
          attempted: semanticOrder.length,
          completed,
          remaining: semanticOrder.length - completed,
          failureIndex: stepIndex,
          error: `Word mantuvo pendiente el tracked change ${resolvedTrackedChangeType} después del paso ${stepIndex + 1} del replace.`,
        },
        recoveryAttempted,
        recoverySucceeded,
      };
    }

    const remainingTrackedChange = this.findTrackedChangeByType(
      reobserved.observation.trackedChanges,
      remainingTrackedChangeType,
    );
    if (remainingTrackedChange) {
      return { observation: reobserved.observation };
    }

    if (reobserved.observation.trackedChanges.length === 0) {
      return {
        observation: reobserved.observation,
        executionReport: {
          attempted: semanticOrder.length,
          completed: semanticOrder.length,
          remaining: 0,
        },
        recoveryAttempted,
        recoverySucceeded,
      };
    }

    return {
      observation: reobserved.observation,
      executionReport: {
        attempted: semanticOrder.length,
        completed,
        remaining: semanticOrder.length - completed,
        failureIndex: stepIndex + 1,
        error:
          "Word reexpuso tracked changes incompatibles con el lado restante del replace.",
      },
      recoveryAttempted,
      recoverySucceeded,
    };
  }

  /** Returns the final replace-step error payload after recovery. */
  private buildReplaceRecoveryError(
    stillPendingTrackedChange: Word.TrackedChange | null,
    trackedChangesCount: number,
    recoveryError?: string,
  ): { error?: string } {
    if (stillPendingTrackedChange === null && trackedChangesCount === 0) {
      return {};
    }

    if (stillPendingTrackedChange) {
      return { error: recoveryError };
    }

    return {
      error:
        "Word dejó el replace en un estado parcial; se cancela cleanup para evitar falso success.",
    };
  }

  /** Re-checks post-execute replace state and retries one atomic accept batch if Word still exposes a full pair. */
  private async observePostExecuteResolutionState(
    context: Word.RequestContext,
    preferredCc: Word.ContentControl,
  ): Promise<{
    observation: ResolutionObservation | null;
    recoveryAttempted: boolean;
    recoverySucceeded: boolean;
  }> {
    const postExecuteObservation = await this.logWorkflowSnapshot(
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

    const normalizedObservation = this.normalizeReplaceObservation(
      postExecuteObservation,
    );
    const hasDeleted =
      this.findTrackedChangeByType(
        normalizedObservation.trackedChanges,
        "Deleted",
      ) !== null;
    const hasAdded =
      this.findTrackedChangeByType(
        normalizedObservation.trackedChanges,
        "Added",
      ) !== null;
    // Ignore adjacent neighbor TCs: only trigger the atomic retry when this
    // suggestion's own CC scope (cc or ccRange) still exposes tracked changes.
    // `bodyRelated` can pull in `AdjacentBefore/AdjacentAfter` TCs from nearby
    // suggestions, which wrongly looks like an unresolved replace pair here.
    const ccScopedRemaining =
      (normalizedObservation.debugMetadata?.ccTrackedChangesCount ?? 0) +
      (normalizedObservation.debugMetadata?.ccRangeTrackedChangesCount ?? 0);

    if (
      this.action !== "accept" ||
      !this.isReplaceSuggestion() ||
      !hasDeleted ||
      !hasAdded ||
      ccScopedRemaining === 0
    ) {
      return {
        observation: postExecuteObservation,
        recoveryAttempted: false,
        recoverySucceeded: false,
      };
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" retrying accept replace as one atomic batch after post-execute snapshot still exposed the full replace pair`,
    );

    const atomicReport = await this.executor.applyAtomically(
      context,
      normalizedObservation.trackedChanges,
    );
    if (atomicReport.error) {
      const semanticPairRecovery =
        await this.tryFreshPostExecuteAcceptSemanticPairRecovery(
          context,
          normalizedObservation.selectedCc,
          atomicReport.error,
          postExecuteObservation,
        );
      if (semanticPairRecovery) {
        return semanticPairRecovery;
      }

      return {
        observation: postExecuteObservation,
        recoveryAttempted: true,
        recoverySucceeded: false,
      };
    }

    const recoveredObservation = await this.logWorkflowSnapshot(
      context,
      "after-post-execute-recovery-before-cleanup",
      normalizedObservation.selectedCc,
    );

    return {
      observation: recoveredObservation ?? postExecuteObservation,
      recoveryAttempted: true,
      recoverySucceeded:
        recoveredObservation !== null &&
        recoveredObservation.trackedChanges.length === 0,
    };
  }

  /** Re-runs one fresh semantic Deleted/Added pass after a stale post-execute atomic retry fails with ItemNotFound. */
  private async tryFreshPostExecuteAcceptSemanticPairRecovery(
    context: Word.RequestContext,
    preferredCc: Word.ContentControl,
    atomicError: string,
    fallbackObservation: ResolutionObservation,
  ): Promise<{
    observation: ResolutionObservation | null;
    recoveryAttempted: boolean;
    recoverySucceeded: boolean;
  } | null> {
    if (!atomicError.includes("ItemNotFound")) {
      return null;
    }

    const deletedReobservation =
      await this.reobserveResolutionCandidatesForSemanticSide(
        context,
        "Deleted",
        preferredCc,
      );
    console.log(
      `🔁 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" post-execute semantic recovery side=Deleted status=${deletedReobservation?.observation.observationStatus ?? "missing"} trackedChanges=${deletedReobservation?.observation.trackedChanges.length ?? 0} types=${deletedReobservation?.observation.debugMetadata?.trackedChangeTypes ?? ""} semanticSource=${deletedReobservation?.observation.debugMetadata?.selectedSemanticSideSource ?? ""}`,
    );
    const deletedTrackedChange =
      deletedReobservation === null
        ? null
        : this.findTrackedChangeByType(
            deletedReobservation.observation.trackedChanges,
            "Deleted",
          );

    if (!deletedTrackedChange) {
      return null;
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" retrying accept replace with a fresh semantic Deleted pass after post-execute atomic failure: ${atomicError}`,
    );

    const deletedRecoveryReport = await this.executor.apply(context, [
      deletedTrackedChange,
    ]);
    if (deletedRecoveryReport.error) {
      return {
        observation: fallbackObservation,
        recoveryAttempted: true,
        recoverySucceeded: false,
      };
    }

    const addedReobservation =
      await this.reobserveResolutionCandidatesForSemanticSide(
        context,
        "Added",
        deletedReobservation.observation.selectedCc,
      );
    console.log(
      `🔁 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" post-execute semantic recovery side=Added status=${addedReobservation?.observation.observationStatus ?? "missing"} trackedChanges=${addedReobservation?.observation.trackedChanges.length ?? 0} types=${addedReobservation?.observation.debugMetadata?.trackedChangeTypes ?? ""} semanticSource=${addedReobservation?.observation.debugMetadata?.selectedSemanticSideSource ?? ""}`,
    );
    const addedTrackedChange =
      addedReobservation === null
        ? null
        : this.findTrackedChangeByType(
            addedReobservation.observation.trackedChanges,
            "Added",
          );

    if (!addedTrackedChange) {
      const addedObservation =
        addedReobservation?.observation ?? fallbackObservation;

      return {
        observation: addedObservation,
        recoveryAttempted: true,
        recoverySucceeded: addedObservation.trackedChanges.length === 0,
      };
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" retrying accept replace with a fresh semantic Added pass after post-execute atomic failure: ${atomicError}`,
    );

    const addedRecoveryReport = await this.executor.apply(context, [
      addedTrackedChange,
    ]);
    if (addedRecoveryReport.error) {
      return {
        observation: addedReobservation.observation,
        recoveryAttempted: true,
        recoverySucceeded: false,
      };
    }

    const recoveredObservation = await this.logWorkflowSnapshot(
      context,
      "after-post-execute-recovery-before-cleanup",
      addedReobservation.observation.selectedCc,
    );

    return {
      observation: recoveredObservation ?? fallbackObservation,
      recoveryAttempted: true,
      recoverySucceeded:
        recoveredObservation !== null &&
        recoveredObservation.trackedChanges.length === 0,
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

  /** Builds one stable tracked-change entry for cross-phase workflow logs. */
  private describeTrackedChangeForLog(
    trackedChange: Word.TrackedChange,
  ): TrackedChangeLogEntry {
    return {
      id: String((trackedChange as { id?: string | number }).id ?? "no-id"),
      type: trackedChange.type ?? "unknown",
    };
  }

  /** Builds a compact tracked-change list so one workflow attempt can be reconstructed later. */
  private describeTrackedChangesForLog(
    trackedChanges: Word.TrackedChange[],
  ): TrackedChangeLogEntry[] {
    return trackedChanges.map((trackedChange) =>
      this.describeTrackedChangeForLog(trackedChange),
    );
  }

  /** Captures a best-effort fresh snapshot around execute/cleanup so false-success runs leave enough host evidence. */
  private async logWorkflowSnapshot(
    context: Word.RequestContext,
    label:
      | "after-execute-before-cleanup"
      | "after-post-execute-recovery-before-cleanup"
      | "after-cleanup-before-return",
    preferredCc?: Word.ContentControl,
    reviewState?: import("../../domain/types").DocumentReviewState,
  ): Promise<ResolutionObservation | null> {
    try {
      const currentReviewState =
        reviewState ?? (await this.stateInspector.inspect(context));

      if (!this.isReplaceSuggestion()) {
        console.log(
          `🧪 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" snapshot=${label}`,
          {
            reviewState: currentReviewState,
            replaceSuggestion: false,
          },
        );
        return null;
      }

      const relocated = await this.locator.locateResolutionArtifacts(context);
      if (!relocated.selectedCc) {
        console.log(
          `🧪 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" snapshot=${label}`,
          {
            reviewState: currentReviewState,
            relocatedCandidateCount: relocated.rankedCandidates.length,
            relocatedSelectedCc: null,
          },
        );
        return null;
      }

      const resolvedPreferredCc = this.resolveFreshPreferredCandidate(
        relocated.rankedCandidates,
        preferredCc,
      );
      const preferredCandidates = this.prioritizeFreshPreferredCandidate(
        relocated.rankedCandidates,
        resolvedPreferredCc,
      );
      const observation = await this.observer.observeResolutionCandidates(
        context,
        preferredCandidates,
        resolvedPreferredCc ?? relocated.selectedCc,
      );

      console.log(
        `🧪 [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" snapshot=${label}`,
        {
          reviewState: currentReviewState,
          relocatedCandidateCount: relocated.rankedCandidates.length,
          relocatedSelectedCc: relocated.selectedCc.tag,
          preferredCcResolved: resolvedPreferredCc?.tag ?? null,
          observationStatus: observation.observationStatus,
          trackedChangesObserved: observation.trackedChanges.length,
          trackedChanges: this.describeTrackedChangesForLog(
            observation.trackedChanges,
          ),
          debugMetadata: observation.debugMetadata ?? null,
        },
      );
      return observation;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" snapshot=${label} failed: ${message}`,
      );
      return null;
    }
  }

  /** Emits best-effort structured telemetry without changing semantic outcomes. */
  private async emitTelemetry(
    phase: ResolutionPhase,
    outcome: ResolutionTelemetryEvent["outcome"],
    metadata?: ResolutionTelemetryEvent["metadata"],
  ): Promise<void> {
    try {
      await this.telemetryPort.emit({
        workflowAttemptId: this.workflowAttemptId,
        suggestionId: this.suggestion.id,
        action: this.action,
        phase,
        outcome,
        metadata,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [ResolveSuggestionCommand] telemetry failed for suggestionId="${this.suggestion.id}" phase=${phase}: ${message}`,
      );
    }
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
