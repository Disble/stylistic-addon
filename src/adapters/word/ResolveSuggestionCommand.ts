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

type SerializedOfficeErrorDiagnostics = {
  message: string;
  name?: string;
  code?: string | number;
  debugInfo?: unknown;
  traceMessages?: unknown;
  stackPreview?: string[];
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

  /** Reads one unknown error property defensively so diagnostic logging never throws. */
  private readUnknownErrorProperty(
    error: unknown,
    propertyName: string,
  ): unknown {
    if (typeof error !== "object" || error === null) {
      return undefined;
    }

    try {
      return (error as Record<string, unknown>)[propertyName];
    } catch {
      return undefined;
    }
  }

  /** Builds one plain Office.js-ish error diagnostic object for console output. */
  private serializeUnknownError(
    error: unknown,
  ): SerializedOfficeErrorDiagnostics {
    const fallbackMessage =
      error instanceof Error ? error.message : String(error ?? "Unknown error");
    const messageValue = this.readUnknownErrorProperty(error, "message");
    const nameValue = this.readUnknownErrorProperty(error, "name");
    const codeValue = this.readUnknownErrorProperty(error, "code");
    const debugInfo = this.readUnknownErrorProperty(error, "debugInfo");
    const traceMessages = this.readUnknownErrorProperty(error, "traceMessages");
    const stackValue = this.readUnknownErrorProperty(error, "stack");
    const stackPreview =
      typeof stackValue === "string"
        ? stackValue
            .split(/\r?\n/u)
            .map((line) => line.trim())
            .filter((line) => line.length > 0)
            .slice(0, 5)
        : undefined;

    return {
      message:
        typeof messageValue === "string" && messageValue.length > 0
          ? messageValue
          : fallbackMessage,
      ...(typeof nameValue === "string" && nameValue.length > 0
        ? { name: nameValue }
        : {}),
      ...(typeof codeValue === "string" || typeof codeValue === "number"
        ? { code: codeValue }
        : {}),
      ...(debugInfo !== undefined ? { debugInfo } : {}),
      ...(traceMessages !== undefined ? { traceMessages } : {}),
      ...(stackPreview && stackPreview.length > 0 ? { stackPreview } : {}),
    };
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
      const serializedError = this.serializeUnknownError(error);
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
    console.log(`🧪 [ResolveSuggestionCommand] observe-before start`, {
      workflowAttemptId: this.workflowAttemptId,
      suggestionId: this.suggestion.id,
      action: this.action,
      selectedCcTag: cc.tag,
      rankedCandidateCount: rankedCandidates.length,
      suggestionType: this.suggestion.type,
    });
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
          error: this.serializeUnknownError(error),
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

  /**
   * Falls back to one atomic accept batch, but only trusts it after a fresh
   * re-observation proves the replace pair is no longer pending.
   */
  private async tryAtomicAcceptReplaceFallback(
    context: Word.RequestContext,
    stepIndex: number,
    initialObservation: ResolutionObservation,
    latestObservation: ResolutionObservation,
    completed: number,
  ): Promise<{
    observation: ResolutionObservation;
    executionReport: ResolutionExecutionReport;
    recoverySucceeded: boolean;
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
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" retrying accept replace as one atomic batch after the first semantic side failed`,
    );

    const atomicReport = await this.executor.applyAtomically(
      context,
      latestObservation.trackedChanges,
    );
    if (atomicReport.error) {
      return null;
    }

    const reobservedAfterAtomic = await this.logWorkflowSnapshot(
      context,
      "after-post-execute-recovery-before-cleanup",
      latestObservation.selectedCc,
    );
    const atomicCompletion = this.classifyAtomicReplaceCompletion(
      reobservedAfterAtomic,
    );

    return {
      observation: reobservedAfterAtomic ?? latestObservation,
      executionReport: atomicCompletion.completed
        ? atomicReport
        : {
            ...atomicReport,
            remaining: atomicCompletion.remaining,
            error:
              "Word siguió exponiendo tracked changes del replace después del fallback atómico; se cancela cleanup para evitar falso success.",
          },
      recoverySucceeded: atomicCompletion.completed,
    };
  }

  /**
   * Tries the atomic accept fallback for one semantic side, but only certifies
   * success after a fresh re-observation proves that side disappeared.
   */
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
      trackedChangeType !== "Added" ||
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
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" retrying accept replace as one atomic batch after Added-alone execution failed`,
    );

    const atomicReport = await this.executor.applyAtomically(
      context,
      observation.trackedChanges,
    );
    if (atomicReport.error) {
      return null;
    }

    const reobservedAfterAtomic = await this.reobserveResolutionCandidates(
      context,
      observation.selectedCc,
    );
    if (!reobservedAfterAtomic) {
      return {
        observation,
        completed: false,
        error:
          "Word no pudo reobservar el replace después del fallback atómico; se cancela cleanup para evitar falso success.",
        recoveryAttempted: true,
        recoverySucceeded: false,
      };
    }

    const sideStillPending =
      this.findTrackedChangeByType(
        reobservedAfterAtomic.observation.trackedChanges,
        trackedChangeType,
      ) !== null;

    return {
      observation: reobservedAfterAtomic.observation,
      completed: !sideStillPending,
      ...(sideStillPending
        ? {
            error: `Word siguió exponiendo el tracked change ${trackedChangeType} después del fallback atómico; se cancela cleanup para evitar falso success.`,
          }
        : {}),
      recoveryAttempted: true,
      recoverySucceeded: !sideStillPending,
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
    if (this.isExecutionReportSemanticallyVerified(initialReport)) {
      return {
        observation,
        completed: true,
        recoveryAttempted: false,
        recoverySucceeded: false,
      };
    }

    let bodyRecoveryError: string | undefined;
    if (initialReport.silentNoOpDetected && !initialReport.error) {
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace-step=${trackedChangeType} silent no-op detected (bodyTrackedChangeCountBefore=${initialReport.silentNoOpDetected.bodyTrackedChangeCountBefore} after=${initialReport.silentNoOpDetected.bodyTrackedChangeCountAfter}); attempting body-text recovery before returning success`,
      );
      const bodyRecovery = await this.recoverFromSilentNoOpForReplaceSide(
        context,
        trackedChangeType,
      );
      if (bodyRecovery.completed) {
        return {
          observation,
          completed: true,
          recoveryAttempted: true,
          recoverySucceeded: true,
        };
      }

      // Body-text recovery did not succeed; surface this as a step error so
      // the existing recovery cascade gets a chance to relocate fresh proxies.
      bodyRecoveryError = bodyRecovery.error;
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace-step=${trackedChangeType} silent-no-op body-text recovery failed: ${bodyRecovery.error ?? "no matching body tracked-change"}`,
      );
    }

    if (initialReport.unverifiedMutation && !initialReport.error) {
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace-step=${trackedChangeType} mutation verification unavailable after ${this.action}; re-observing fresh Word state before returning success`,
        initialReport.unverifiedMutation,
      );
    }

    const silentNoOpSuffix = bodyRecoveryError
      ? ` [recovery: ${bodyRecoveryError}]`
      : "";
    const initialUnverifiedMutation = initialReport.unverifiedMutation;
    const initialErrorMessage =
      initialReport.error ??
      (initialUnverifiedMutation
        ? `Word no pudo verificar si el ${this.action === "reject" ? "rechazo" : "aceptación"} del lado ${trackedChangeType} mutó el documento (${this.formatUnverifiedMutationForLog(initialUnverifiedMutation)}).${silentNoOpSuffix}`
        : `Word ignoró el ${this.action === "reject" ? "rechazo" : "aceptación"} del lado ${trackedChangeType} (silent no-op detectado: el proxy del tracked change no mutó el documento).${silentNoOpSuffix}`);

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace-step=${trackedChangeType} retrying after failure: ${initialErrorMessage}`,
    );

    const firstRecoveryObservation = await this.reobserveResolutionCandidates(
      context,
      observation.selectedCc,
    );
    if (!firstRecoveryObservation) {
      return {
        observation,
        completed: false,
        error: initialErrorMessage,
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
      return {
        observation: firstRecoveryObservation.observation,
        completed: true,
        recoveryAttempted: true,
        recoverySucceeded: true,
      };
    }

    const recoveryReport = await this.executor.apply(context, [
      recoveredTrackedChange,
    ]);
    if (this.isExecutionReportSemanticallyVerified(recoveryReport)) {
      return {
        observation: firstRecoveryObservation.observation,
        completed: true,
        recoveryAttempted: true,
        recoverySucceeded: true,
      };
    }

    if (recoveryReport.silentNoOpDetected) {
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace-step=${trackedChangeType} recovered proxy was a silent no-op; validating with final re-observation`,
        recoveryReport.silentNoOpDetected,
      );
    }

    if (recoveryReport.unverifiedMutation) {
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace-step=${trackedChangeType} recovered proxy mutation verification unavailable; validating with final re-observation`,
        recoveryReport.unverifiedMutation,
      );
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
    const recoveredSideCompleted = stillPendingTrackedChange === null;

    return {
      observation: finalRecoveryObservation.observation,
      completed: recoveredSideCompleted,
      ...(recoveredSideCompleted
        ? {}
        : {
            error:
              recoveryReport.error ??
              this.buildUntrustedExecutionError(
                recoveryReport,
                trackedChangeType,
              ),
          }),
      recoveryAttempted: true,
      recoverySucceeded: recoveredSideCompleted,
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

  /**
   * Recovers from a silent no-op resolution by reaching for a fresh tracked-change
   * proxy from `body.getTrackedChanges()` whose range text matches this
   * suggestion's expected side text.
   *
   * The silent no-op pattern happens in real Word when the executor receives a
   * stale `Word.TrackedChange` proxy obtained from `ccRange.getTrackedChanges()`
   * (or `cc.getTrackedChanges()`) — the host-side `accept()/reject()` resolves
   * cleanly and `context.sync()` returns no error, but the document is not
   * mutated. The body-level `getTrackedChanges()` call returns a different
   * proxy backed by the actual document range; reapplying the action on that
   * proxy mutates the document for real.
   *
   * Returns `{ completed: true }` only if the body-text recovery actually
   * decreased the document tracked-change count (i.e., the recovery proxy was
   * not silent too). Otherwise returns `{ completed: false, error }` so the
   * caller can decide whether to surface the failure or fall back to the
   * existing reobserve-and-retry cascade.
   *
   * Restricted by text matching to avoid accidentally resolving tracked
   * changes that belong to a neighboring suggestion.
   */
  private async recoverFromSilentNoOpForReplaceSide(
    context: Word.RequestContext,
    trackedChangeType: "Added" | "Deleted",
  ): Promise<{ completed: boolean; error?: string }> {
    const expectedText =
      trackedChangeType === "Deleted"
        ? this.suggestion.anchor
        : (this.suggestion.suggestedText ?? "");

    if (expectedText.length === 0) {
      return {
        completed: false,
        error: `silent-no-op recovery aborted: missing expected text for ${trackedChangeType} side`,
      };
    }

    let bodyTrackedChangesBefore = 0;
    let candidateRanges: Word.Range[] = [];
    let candidates: Word.TrackedChange[] = [];

    try {
      const body = context.document.body;
      const bodyTrackedChanges = body.getTrackedChanges();
      bodyTrackedChanges.load({ select: "type,id" });
      await context.sync();
      bodyTrackedChangesBefore = bodyTrackedChanges.items.length;

      candidates = bodyTrackedChanges.items.filter(
        (tc) => tc.type === trackedChangeType,
      );
      if (candidates.length === 0) {
        return {
          completed: false,
          error: `silent-no-op recovery: body exposed 0 ${trackedChangeType} tracked changes`,
        };
      }

      candidateRanges = candidates.map((tc) => tc.getRange());
      candidateRanges.forEach((range) => {
        range.load({ select: "text" });
      });
      await context.sync();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        completed: false,
        error: `silent-no-op recovery probe failed: ${message}`,
      };
    }

    // Diagnostic: log every candidate so we can refine matching when it fails.
    candidateRanges.forEach((range, index) => {
      const tc = candidates[index];
      const trackedChangeLog = tc
        ? this.describeTrackedChangeForLog(tc)
        : { id: "unknown", type: "unknown" };
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" silent-no-op body candidate[${index}] type=${trackedChangeLog.type} id=${trackedChangeLog.id} text="${(range.text ?? "").slice(0, 120)}"`,
      );
    });

    const normalizeForMatch = (value: string): string =>
      value.trim().replace(/\s+/gu, " ").toLowerCase();
    const normalizedExpected = normalizeForMatch(expectedText);
    let matchingIndex = candidateRanges.findIndex((range) => {
      const rangeText = normalizeForMatch(range.text ?? "");
      if (rangeText.length === 0 || normalizedExpected.length === 0) {
        return false;
      }
      return (
        rangeText === normalizedExpected ||
        rangeText.includes(normalizedExpected) ||
        normalizedExpected.includes(rangeText)
      );
    });

    // Single-candidate fallback: when text matching fails but only ONE
    // body tracked-change of the requested type exists, it must be the one
    // we wanted to resolve. Tracked-change ranges sometimes expose only the
    // edited delta or stale text after the sibling side was resolved, so
    // text matching alone is unreliable. We still gate this fallback on
    // having exactly one candidate to avoid touching neighboring suggestions.
    if (matchingIndex === -1 && candidates.length === 1) {
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" silent-no-op single-candidate fallback for ${trackedChangeType} (text matching failed but only one body candidate of this type exists)`,
      );
      matchingIndex = 0;
    }

    if (matchingIndex === -1) {
      return {
        completed: false,
        error: `silent-no-op recovery: no body ${trackedChangeType} tracked-change matched expected text "${expectedText.slice(0, 80)}" among ${candidates.length} candidates`,
      };
    }

    const matchingTrackedChange = candidates[matchingIndex];
    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" silent-no-op body-text recovery applying ${this.action} on body ${trackedChangeType} tracked-change matching expected text`,
    );

    const recoveryReport = await this.executor.apply(context, [
      matchingTrackedChange,
    ]);
    if (recoveryReport.error) {
      return {
        completed: false,
        error: `silent-no-op recovery apply failed: ${recoveryReport.error}`,
      };
    }
    if (recoveryReport.silentNoOpDetected) {
      return {
        completed: false,
        error: `silent-no-op recovery: body proxy was also a silent no-op (bodyTrackedChangeCountBefore=${recoveryReport.silentNoOpDetected.bodyTrackedChangeCountBefore} after=${recoveryReport.silentNoOpDetected.bodyTrackedChangeCountAfter})`,
      };
    }
    if (recoveryReport.unverifiedMutation) {
      return {
        completed: false,
        error: `silent-no-op recovery: body proxy mutation could not be verified (${this.formatUnverifiedMutationForLog(recoveryReport.unverifiedMutation)})`,
      };
    }

    // Confirm document tracked-change count actually decreased.
    let bodyTrackedChangesAfter = bodyTrackedChangesBefore;
    try {
      const body = context.document.body;
      const bodyTrackedChanges = body.getTrackedChanges();
      bodyTrackedChanges.load({ select: "type,id" });
      await context.sync();
      bodyTrackedChangesAfter = bodyTrackedChanges.items.length;
    } catch {
      // If we can't probe the count again, trust the executor's own check.
    }

    if (bodyTrackedChangesAfter >= bodyTrackedChangesBefore) {
      return {
        completed: false,
        error: `silent-no-op recovery: bodyTrackedChangeCount did not decrease (before=${bodyTrackedChangesBefore} after=${bodyTrackedChangesAfter})`,
      };
    }

    return { completed: true };
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

  /** Returns true when the current suggestion is a tracked replace. */
  private isReplaceSuggestion(): boolean {
    return (
      this.suggestion.type === "track-change" &&
      this.suggestion.anchor.length > 0 &&
      (this.suggestion.suggestedText?.length ?? 0) > 0
    );
  }

  /** Returns true only when execution has no error and no unknown host-verification state. */
  private isExecutionReportSemanticallyVerified(
    report: ResolutionExecutionReport,
  ): boolean {
    return (
      !report.error && !report.silentNoOpDetected && !report.unverifiedMutation
    );
  }

  /** Formats an unverified mutation signal for single-line workflow errors. */
  private formatUnverifiedMutationForLog(
    unverifiedMutation: NonNullable<
      ResolutionExecutionReport["unverifiedMutation"]
    >,
  ): string {
    const before = unverifiedMutation.bodyTrackedChangeCountBefore ?? "unknown";
    const after = unverifiedMutation.bodyTrackedChangeCountAfter ?? "unknown";
    const beforeError =
      unverifiedMutation.bodyTrackedChangeCountBeforeError ?? null;
    const afterError =
      unverifiedMutation.bodyTrackedChangeCountAfterError ?? null;

    return [
      `bodyTrackedChangeCountBefore=${before}`,
      `bodyTrackedChangeCountAfter=${after}`,
      beforeError ? `beforeError=${beforeError}` : null,
      afterError ? `afterError=${afterError}` : null,
    ]
      .filter((part): part is string => part !== null)
      .join("; ");
  }

  /** Builds a conservative error for an execution report that cannot certify mutation. */
  private buildUntrustedExecutionError(
    report: ResolutionExecutionReport,
    trackedChangeType: "Added" | "Deleted",
  ): string {
    if (report.error) {
      return report.error;
    }

    if (report.unverifiedMutation) {
      return `Word no pudo verificar si el ${this.action === "reject" ? "rechazo" : "aceptación"} del lado ${trackedChangeType} mutó el documento (${this.formatUnverifiedMutationForLog(report.unverifiedMutation)}).`;
    }

    if (report.silentNoOpDetected) {
      return `Word ignoró el ${this.action === "reject" ? "rechazo" : "aceptación"} del lado ${trackedChangeType} (silent no-op detectado: el proxy del tracked change no mutó el documento).`;
    }

    return `Word no pudo certificar la resolución del tracked change ${trackedChangeType}.`;
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
      ...((recoveryReport.silentNoOpDetected ??
      initialReport.silentNoOpDetected)
        ? {
            silentNoOpDetected:
              recoveryReport.silentNoOpDetected ??
              initialReport.silentNoOpDetected,
          }
        : {}),
      ...((recoveryReport.unverifiedMutation ??
      initialReport.unverifiedMutation)
        ? {
            unverifiedMutation:
              recoveryReport.unverifiedMutation ??
              initialReport.unverifiedMutation,
          }
        : {}),
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
   * Accept and reject intentionally diverge because real Word exposes different
   * survivable host surfaces per action:
   * - accept: `Added -> Deleted`, so the workflow uses the inserted-side host
   *   surface before stale Deleted proxies can silently no-op.
   * - reject: `Deleted -> Added`, because rejecting the Deleted side first keeps
   *   the inserted-side CC alive long enough to re-observe the remaining Added
   *   side with fresh proxies.
   */
  private getReplaceSemanticOrder():
    | readonly ["Deleted", "Added"]
    | readonly ["Added", "Deleted"] {
    if (this.action === "accept") {
      return ["Added", "Deleted"] as const;
    }

    return ["Deleted", "Added"] as const;
  }

  /** Builds the terminal outcome for a replace-step failure, preserving one-shot atomic fallback. */
  private buildReplaceFailureOutcome(
    atomicFallback: {
      observation: ResolutionObservation;
      executionReport: ResolutionExecutionReport;
      recoverySucceeded: boolean;
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
        recoverySucceeded: atomicFallback.recoverySucceeded,
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

  /**
   * Classifies whether an atomic replace fallback is semantically complete.
   *
   * Success is trusted ONLY after a fresh observation proves there is no
   * remaining CC-scoped confirmed-pending replace pair for this suggestion.
   */
  private classifyAtomicReplaceCompletion(
    observation: ResolutionObservation | null,
  ): {
    completed: boolean;
    remaining: number;
  } {
    if (!observation) {
      return {
        completed: false,
        remaining: 1,
      };
    }

    const ccScopedRemaining =
      (observation.debugMetadata?.ccTrackedChangesCount ?? 0) +
      (observation.debugMetadata?.ccRangeTrackedChangesCount ?? 0);
    const replaceStillPending =
      observation.observationStatus === "confirmed-pending" &&
      ccScopedRemaining > 0;

    return {
      completed: !replaceStillPending,
      remaining: replaceStillPending ? ccScopedRemaining : 0,
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
        this.classifyAtomicReplaceCompletion(recoveredObservation).completed,
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
    if (!this.isExecutionReportSemanticallyVerified(deletedRecoveryReport)) {
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
        recoverySucceeded:
          this.classifyAtomicReplaceCompletion(addedObservation).completed,
      };
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" retrying accept replace with a fresh semantic Added pass after post-execute atomic failure: ${atomicError}`,
    );

    const addedRecoveryReport = await this.executor.apply(context, [
      addedTrackedChange,
    ]);
    if (!this.isExecutionReportSemanticallyVerified(addedRecoveryReport)) {
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
        this.classifyAtomicReplaceCompletion(recoveredObservation).completed,
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
      const serializedError = this.serializeUnknownError(error);
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" snapshot=${label} failed: ${serializedError.message}`,
        {
          error: serializedError,
        },
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
