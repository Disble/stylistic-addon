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
import type {
  ResolutionExecutionReport,
  ResolutionPhase,
  ResolutionTelemetryEvent,
  Suggestion,
  SuggestionActionResult,
  SuggestionResolutionWarning,
} from "../../domain/types";
import { CommentOnlySuggestionResolver } from "./resolution/CommentOnlySuggestionResolver";
import { DocumentReviewStateInspector } from "./resolution/DocumentReviewStateInspector";
import type { ResolutionObservation } from "./resolution/ResolutionContext";
import { ResolveSuggestionResultFactory } from "./resolution/ResolveSuggestionResultFactory";
import { SuggestionLocator } from "./resolution/SuggestionLocator";
import { SuggestionResolutionCleanup } from "./resolution/SuggestionResolutionCleanup";
import { SuggestionResolutionObserver } from "./resolution/SuggestionResolutionObserver";
import { SuggestionResolutionResolver } from "./resolution/SuggestionResolutionResolver";
import { TrackedChangeResolutionExecutor } from "./resolution/TrackedChangeResolutionExecutor";
import {
  getDefaultTextLocator,
  type TextLocator,
} from "./WordTextLocatorContext";

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
  private readonly resolver: SuggestionResolutionResolver;
  private lastExecutionReport?: ResolutionExecutionReport;
  private workflowAttemptId = "";
  private telemetryWarnings: SuggestionResolutionWarning[] = [];

  /** Captures one execute phase, including any same-click recovery attempt. */
  private typeSafeNoop?: never;

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
    this.resolver = new SuggestionResolutionResolver(
      suggestion.id,
      this.resultFactory,
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
        this.lastExecutionReport = undefined;
        this.workflowAttemptId = this.buildWorkflowAttemptId();
        this.telemetryWarnings = [];
        console.log(
          `🎯 [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" type=${this.suggestion.type}`,
        );
        await this.emitTelemetry("observe-before", "started", {
          suggestionType: this.suggestion.type,
        });
        const pendingBefore = await this.stateInspector.inspect(context);
        const { rankedCandidates, selectedCc: cc } =
          await this.locator.locateResolutionArtifacts(context);

        if (!cc) {
          console.warn(
            `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" failed: CC not found`,
          );
          await this.emitTelemetry("observe-before", "failed", {
            reason: "cc-not-found",
          });
          return this.resultFactory.buildResolutionResult(
            "cc-not-found",
            0,
            false,
            pendingBefore,
            pendingBefore,
          );
        }

        console.log(
          `🎯 [ResolveSuggestionCommand] selected CC for suggestionId="${this.suggestion.id}": tag="${cc.tag}"`,
        );

        if (this.suggestion.type === "comment-only") {
          const colocatedComment =
            await this.locator.findColocatedStylisticComment(context, cc);
          const commentDeleted =
            await this.cleanup.deleteLocatedStylisticComment(
              context,
              colocatedComment,
            );
          await this.emitTelemetry("cleanup", "succeeded", {
            commentDeleted,
            commentOnly: true,
          });
          return this.commentOnlyResolver.resolve({
            context,
            cc,
            commentDeleted,
            pendingBefore,
          });
        }

        const observation = await this.observer.observeResolutionCandidates(
          context,
          rankedCandidates,
          cc,
        );

        if (observation.observationStatus === "identity-lost") {
          console.warn(
            `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" ended in identity-lost`,
          );
          await this.emitTelemetry("observe-before", "failed", {
            reason: "identity-lost",
          });
          return this.resultFactory.buildObservationFailureResult(
            context,
            "identity-lost",
            pendingBefore,
          );
        }

        if (
          observation.observationStatus !== "confirmed-pending" ||
          observation.trackedChanges.length === 0
        ) {
          console.warn(
            `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" ended in unobservable after all evidence sources`,
          );
          await this.emitTelemetry("observe-before", "warning", {
            reason: "unobservable",
          });
          return this.resultFactory.buildObservationFailureResult(
            context,
            "unobservable",
            pendingBefore,
          );
        }

        await this.emitTelemetry("observe-before", "succeeded", {
          trackedChangesObserved: observation.trackedChanges.length,
        });
        await this.emitTelemetry("execute", "started", {
          trackedChangesAttempted: observation.trackedChanges.length,
        });
        const executeAttempt =
          await this.executeTrackedChangesWithImmediateRecovery(
            context,
            rankedCandidates,
            observation,
          );
        const executionReport = executeAttempt.executionReport;
        this.lastExecutionReport = executionReport;
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
          },
        );

        if (executionReport.error) {
          throw new Error(executionReport.error);
        }

        const warnings: SuggestionResolutionWarning[] = [
          ...this.telemetryWarnings,
        ];
        const commentCleanup =
          await this.cleanup.deleteLocatedStylisticCommentAfterResolution(
            context,
            executeAttempt.observation.selectedComment,
          );
        if (commentCleanup.warning) {
          warnings.push(commentCleanup.warning);
        }
        await this.emitTelemetry("cleanup", "succeeded", {
          commentDeleted: commentCleanup.deleted,
          commentCleanupWarning: !!commentCleanup.warning,
        });

        const anchorCleanupWarning =
          await this.cleanup.cleanupResolvedSuggestionAnchor(
            context,
            executeAttempt.observation.selectedCc,
          );
        if (anchorCleanupWarning) {
          warnings.push(anchorCleanupWarning);
        }

        const inspectAfter = await this.stateInspector.inspectAfterResolution(
          context,
          pendingBefore,
          this.action,
          this.suggestion.id,
        );
        if (inspectAfter.warning) {
          warnings.push(inspectAfter.warning);
        }
        await this.emitTelemetry("inspect-after", "succeeded", {
          pendingArtifacts: inspectAfter.pendingAfter.pendingStylisticArtifacts,
          inspectionWarning: !!inspectAfter.warning,
        });

        return this.resultFactory.buildResolutionResult(
          this.resultFactory.toResolutionStatus(),
          executionReport.completed,
          commentCleanup.deleted,
          pendingBefore,
          inspectAfter.pendingAfter,
          undefined,
          warnings,
          executionReport,
        );
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [ResolveSuggestionCommand] threw for suggestionId="${this.suggestion.id}": ${message}`,
      );
      const pendingAfter = await Word.run((ctx) =>
        this.stateInspector.inspect(ctx),
      ).catch(() => this.stateInspector.buildEmptyState());

      const reconciliation = await this.tryBuildSemanticReconciliationResult(
        pendingAfter,
        message,
      );
      if (reconciliation) {
        await this.emitTelemetry("reconcile", "reconciled", {
          status: reconciliation.status,
          warnings: reconciliation.warnings?.length ?? 0,
        });
        return reconciliation;
      }

      await this.emitTelemetry("reconcile", "failed", {
        reason: message,
        pendingArtifacts: pendingAfter.pendingStylisticArtifacts,
      });

      return this.resultFactory.buildErrorResult(
        message,
        pendingAfter,
        undefined,
        this.lastExecutionReport,
      );
    }
  }

  /** Reconciles late failures against the semantic execution report. */
  private async tryBuildSemanticReconciliationResult(
    pendingAfter: import("../../domain/types").DocumentReviewState,
    message: string,
  ): Promise<SuggestionActionResult | null> {
    return this.resolver.reconcileLateFailure(
      pendingAfter,
      this.lastExecutionReport,
      message,
      this.telemetryWarnings,
    );
  }

  /** Executes tracked changes and retries once after partial progress if re-observation can recover the remainder. */
  private async executeTrackedChangesWithImmediateRecovery(
    context: Word.RequestContext,
    rankedCandidates: Word.ContentControl[],
    observation: ResolutionObservation,
  ): Promise<{
    observation: ResolutionObservation;
    executionReport: ResolutionExecutionReport;
    recoveryAttempted: boolean;
    recoverySucceeded: boolean;
  }> {
    const initialReport = this.executor.apply(observation.trackedChanges);

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

      if (
        recoveryObservation.observationStatus !== "confirmed-pending" ||
        recoveryObservation.trackedChanges.length === 0
      ) {
        return {
          observation,
          executionReport: initialReport,
          recoveryAttempted: true,
          recoverySucceeded: false,
        };
      }

      const recoveryReport = this.executor.apply(
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

  /** Merges the initial partial execution report with one immediate recovery pass. */
  private mergeExecutionReports(
    initialReport: ResolutionExecutionReport,
    recoveryReport: ResolutionExecutionReport,
  ): ResolutionExecutionReport {
    const completed = initialReport.completed + recoveryReport.completed;
    const remaining = recoveryReport.remaining;

    return {
      attempted: completed + remaining,
      completed,
      remaining,
      ...(recoveryReport.failureIndex !== undefined
        ? {
            failureIndex: initialReport.completed + recoveryReport.failureIndex,
          }
        : {}),
      ...(recoveryReport.error ? { error: recoveryReport.error } : {}),
    };
  }

  /** Builds a correlation id for one resolution workflow attempt. */
  private buildWorkflowAttemptId(): string {
    return `${this.suggestion.id}:${this.action}:${Date.now()}`;
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
      this.telemetryWarnings.push({
        code: "telemetry-failed",
        phase,
        message,
      });
      console.warn(
        `⚠️ [ResolveSuggestionCommand] telemetry failed for suggestionId="${this.suggestion.id}" phase=${phase}: ${message}`,
      );
    }
  }
}
