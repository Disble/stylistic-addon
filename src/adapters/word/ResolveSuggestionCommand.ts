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
import type {
  ReplaceResolutionStrategy,
  ReplaceTrackedChangeSide,
} from "./resolution/ReplaceResolutionStrategyContext";
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
  private readonly replaceResolutionStrategy: ReplaceResolutionStrategy;
  private readonly resultFactory: ResolveSuggestionResultFactory;
  private readonly commentOnlyResolver: CommentOnlySuggestionResolver;
  private readonly observer: SuggestionResolutionObserver;
  private readonly executeStateMachine = new ExecuteResolutionStateMachine();
  private lastExecutionReport?: ResolutionExecutionReport;
  private workflowAttemptId = "";

  constructor(
    private readonly suggestion: Suggestion,
    private readonly action: "accept" | "reject",
    replaceResolutionStrategy: ReplaceResolutionStrategy,
    textLocator: TextLocator = getDefaultTextLocator(),
    private readonly telemetryPort: ITelemetryPort = {
      emit: async () => undefined,
    },
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

  /** Converts one unknown Office.js-ish error into a stable message without stringifying opaque objects. */
  private stringifyUnknownError(error: unknown): string {
    if (error instanceof Error) {
      return error.message;
    }

    if (typeof error === "string") {
      return error;
    }

    if (
      typeof error === "number" ||
      typeof error === "boolean" ||
      typeof error === "bigint"
    ) {
      return String(error);
    }

    const messageValue = this.readUnknownErrorProperty(error, "message");
    if (typeof messageValue === "string" && messageValue.trim().length > 0) {
      return messageValue;
    }

    return "Unknown error";
  }

  /** Builds one plain Office.js-ish error diagnostic object for console output. */
  private serializeUnknownError(
    error: unknown,
  ): SerializedOfficeErrorDiagnostics {
    const fallbackMessage = this.stringifyUnknownError(error);
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

    const serializedError: SerializedOfficeErrorDiagnostics = {
      message:
        typeof messageValue === "string" && messageValue.length > 0
          ? messageValue
          : fallbackMessage,
    };

    if (typeof nameValue === "string" && nameValue.length > 0) {
      serializedError.name = nameValue;
    }

    if (typeof codeValue === "string" || typeof codeValue === "number") {
      serializedError.code = codeValue;
    }

    if (debugInfo !== undefined) {
      serializedError.debugInfo = debugInfo;
    }

    if (traceMessages !== undefined) {
      serializedError.traceMessages = traceMessages;
    }

    if (stackPreview && stackPreview.length > 0) {
      serializedError.stackPreview = stackPreview;
    }

    return serializedError;
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

    const executeAttempt = await this.executeTrackedChangesWithFreshProxyRetry(
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

  /** Executes tracked changes and retries once with fresh proxies when certification fails. */
  private async executeTrackedChangesWithFreshProxyRetry(
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
      `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" retrying with fresh proxies after uncertified execute result: ${initialReport.error}`,
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
        `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" fresh-proxy retry failed: ${message}`,
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

    const semanticOrder = this.replaceResolutionStrategy.semanticOrder;
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
        return {
          observation: activeObservation,
          executionReport: {
            attempted: semanticOrder.length,
            completed,
            remaining: semanticOrder.length - completed,
            failureIndex: stepIndex,
            error: stepResult.error,
          },
          recoveryAttempted,
          recoverySucceeded,
        };
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

    this.logUnverifiedReplaceSemanticStep(trackedChangeType, initialReport);
    const initialErrorMessage = this.buildReplaceSemanticStepErrorMessage(
      trackedChangeType,
      initialReport,
      "",
    );

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

  /** Builds one stable semantic-step error message after Word failed to certify mutation. */
  private buildReplaceSemanticStepErrorMessage(
    trackedChangeType: ReplaceTrackedChangeSide,
    report: ResolutionExecutionReport,
    silentNoOpSuffix: string,
  ): string {
    if (report.error) {
      return report.error;
    }

    const actionLabel = this.replaceResolutionStrategy.actionLabel;
    if (report.unverifiedMutation) {
      return `Word no pudo verificar si el ${actionLabel} del lado ${trackedChangeType} mutó el documento (${this.formatUnverifiedMutationForLog(report.unverifiedMutation)}).${silentNoOpSuffix}`;
    }

    return `Word ignoró el ${actionLabel} del lado ${trackedChangeType} (silent no-op detectado: el proxy del tracked change no mutó el documento).${silentNoOpSuffix}`;
  }

  /** Logs when Word mutated a replace side but the body-count probe could not certify it. */
  private logUnverifiedReplaceSemanticStep(
    trackedChangeType: "Added" | "Deleted",
    report: ResolutionExecutionReport,
  ): void {
    if (!report.unverifiedMutation || report.error) {
      return;
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${this.workflowAttemptId}" replace-step=${trackedChangeType} mutation verification unavailable after ${this.action}; re-observing fresh Word state before returning success`,
      report.unverifiedMutation,
    );
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
    trackedChangeType: ReplaceTrackedChangeSide,
  ): string {
    if (report.error) {
      return report.error;
    }

    if (report.unverifiedMutation) {
      return `Word no pudo verificar si el ${this.replaceResolutionStrategy.actionLabel} del lado ${trackedChangeType} mutó el documento (${this.formatUnverifiedMutationForLog(report.unverifiedMutation)}).`;
    }

    if (report.silentNoOpDetected) {
      return `Word ignoró el ${this.replaceResolutionStrategy.actionLabel} del lado ${trackedChangeType} (silent no-op detectado: el proxy del tracked change no mutó el documento).`;
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

  /** Re-observes only the remaining replace side and rejects any reappearance of the resolved side. */
  private async reobserveRemainingReplaceSide(
    context: Word.RequestContext,
    activeObservation: ResolutionObservation,
    semanticOrder: readonly [
      ReplaceTrackedChangeSide,
      ReplaceTrackedChangeSide,
    ],
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

  /** Re-checks post-execute replace state without executing any fallback recovery. */
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
    label: "after-execute-before-cleanup" | "after-cleanup-before-return",
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
