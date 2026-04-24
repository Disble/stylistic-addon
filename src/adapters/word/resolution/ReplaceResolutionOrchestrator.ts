/* global console */

import type { ResolutionExecutionReport } from "../../../domain/types";
import type {
  ReplaceResolutionStrategy,
  ReplaceTrackedChangeSide,
} from "./ReplaceResolutionStrategyContext";
import type { ResolutionObservation } from "./ResolutionContext";
import {
  findTrackedChangeByType,
  prioritizeFreshPreferredCandidate,
  resolveFreshPreferredCandidate,
} from "./ResolutionObservationContext";
import type { SuggestionLocator } from "./SuggestionLocator";
import type { SuggestionResolutionObserver } from "./SuggestionResolutionObserver";
import type { TrackedChangeResolutionExecutor } from "./TrackedChangeResolutionExecutor";

/** Result of one replace-resolution attempt with semantic recovery details. */
export type ReplaceResolutionAttempt = {
  observation: ResolutionObservation;
  executionReport: ResolutionExecutionReport;
  recoveryAttempted: boolean;
  recoverySucceeded: boolean;
};

type ReplaceProgressState = {
  completed: number;
  recoveryAttempted: boolean;
  recoverySucceeded: boolean;
};

/** Executes replace resolutions as a dedicated semantic workflow instead of embedding the flow inside the command. */
export class ReplaceResolutionWorkflow {
  constructor(
    private readonly action: "accept" | "reject",
    private readonly locator: SuggestionLocator,
    private readonly observer: SuggestionResolutionObserver,
    private readonly executor: TrackedChangeResolutionExecutor,
    private readonly replaceResolutionStrategy: ReplaceResolutionStrategy,
  ) {}

  /** Resolves replace suggestions in semantic host steps with fresh re-observation between them. */
  async execute(
    context: Word.RequestContext,
    observation: ResolutionObservation,
    workflowAttemptId: string,
  ): Promise<ReplaceResolutionAttempt> {
    if (this.hasDuplicateReplaceSide(observation.trackedChanges)) {
      const normalizedObservation =
        this.normalizeReplaceObservation(observation);
      console.log(
        `⚙️ [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" duplicate-side replace normalized to semantic pair types=${normalizedObservation.debugMetadata?.trackedChangeTypes ?? ""}`,
      );
      return this.execute(context, normalizedObservation, workflowAttemptId);
    }

    const semanticOrder = this.replaceResolutionStrategy.semanticOrder;
    let activeObservation = observation;
    let completed = 0;
    let recoveryAttempted = false;
    let recoverySucceeded = false;

    for (const [stepIndex, trackedChangeType] of semanticOrder.entries()) {
      const stepResult = await this.executeSemanticStep(
        context,
        activeObservation,
        trackedChangeType,
        workflowAttemptId,
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
          {
            completed,
            recoveryAttempted,
            recoverySucceeded,
          },
          workflowAttemptId,
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
  private async executeSemanticStep(
    context: Word.RequestContext,
    observation: ResolutionObservation,
    trackedChangeType: ReplaceTrackedChangeSide,
    workflowAttemptId: string,
  ): Promise<{
    observation: ResolutionObservation;
    completed: boolean;
    error?: string;
    recoveryAttempted: boolean;
    recoverySucceeded: boolean;
  }> {
    const trackedChange = findTrackedChangeByType(
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

    this.logUnverifiedReplaceSemanticStep(
      trackedChangeType,
      initialReport,
      workflowAttemptId,
    );
    const initialErrorMessage = this.buildReplaceSemanticStepErrorMessage(
      trackedChangeType,
      initialReport,
      "",
    );

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace-step=${trackedChangeType} retrying after failure: ${initialErrorMessage}`,
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
      `🔁 [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace-step=${trackedChangeType} recovery observation status=${firstRecoveryObservation.observation.observationStatus} trackedChanges=${firstRecoveryObservation.observation.trackedChanges.length} types=${firstRecoveryObservation.observation.debugMetadata?.trackedChangeTypes ?? ""}`,
    );

    const recoveredTrackedChange = findTrackedChangeByType(
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
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace-step=${trackedChangeType} recovered proxy was a silent no-op; validating with final re-observation`,
        recoveryReport.silentNoOpDetected,
      );
    }

    if (recoveryReport.unverifiedMutation) {
      console.warn(
        `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace-step=${trackedChangeType} recovered proxy mutation verification unavailable; validating with final re-observation`,
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
      `🔁 [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace-step=${trackedChangeType} final recovery observation status=${finalRecoveryObservation.observation.observationStatus} trackedChanges=${finalRecoveryObservation.observation.trackedChanges.length} types=${finalRecoveryObservation.observation.debugMetadata?.trackedChangeTypes ?? ""}`,
    );

    const stillPendingTrackedChange = findTrackedChangeByType(
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

    const resolvedPreferredCc = resolveFreshPreferredCandidate(
      relocated.rankedCandidates,
      preferredCc,
    );
    const preferredCandidates = prioritizeFreshPreferredCandidate(
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
    trackedChangeType: ReplaceTrackedChangeSide,
    preferredCc?: Word.ContentControl,
  ): Promise<{
    rankedCandidates: Word.ContentControl[];
    observation: ResolutionObservation;
  } | null> {
    const relocated = await this.locator.locateResolutionArtifacts(context);
    if (!relocated.selectedCc) {
      return null;
    }

    const resolvedPreferredCc = resolveFreshPreferredCandidate(
      relocated.rankedCandidates,
      preferredCc,
    );
    const preferredCandidates = prioritizeFreshPreferredCandidate(
      relocated.rankedCandidates,
      resolvedPreferredCc,
    );

    const reobserved =
      await this.observer.observeResolutionCandidatesForSemanticSide(
        context,
        preferredCandidates,
        resolvedPreferredCc ?? relocated.selectedCc,
        trackedChangeType,
      );

    return {
      rankedCandidates: relocated.rankedCandidates,
      observation: reobserved,
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
    progress: ReplaceProgressState,
    workflowAttemptId: string,
  ): Promise<
    | {
        observation: ResolutionObservation;
      }
    | ReplaceResolutionAttempt
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
          completed: progress.completed,
          remaining: semanticOrder.length - progress.completed,
          failureIndex: stepIndex + 1,
          error:
            "Word no pudo reubicar la sugerencia después del primer paso del replace.",
        },
        recoveryAttempted: progress.recoveryAttempted,
        recoverySucceeded: progress.recoverySucceeded,
      };
    }

    console.log(
      `🔁 [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace re-observation step=${stepIndex + 1} side=${remainingTrackedChangeType} status=${reobserved.observation.observationStatus} trackedChanges=${reobserved.observation.trackedChanges.length} types=${reobserved.observation.debugMetadata?.trackedChangeTypes ?? ""} semanticSource=${reobserved.observation.debugMetadata?.selectedSemanticSideSource ?? ""} deletedSource=${reobserved.observation.debugMetadata?.selectedDeletedSource ?? ""} addedSource=${reobserved.observation.debugMetadata?.selectedAddedSource ?? ""}`,
    );

    const resolvedSideStillPending = findTrackedChangeByType(
      reobserved.observation.trackedChanges,
      resolvedTrackedChangeType,
    );
    if (resolvedSideStillPending) {
      return {
        observation: reobserved.observation,
        executionReport: {
          attempted: semanticOrder.length,
          completed: progress.completed,
          remaining: semanticOrder.length - progress.completed,
          failureIndex: stepIndex,
          error: `Word mantuvo pendiente el tracked change ${resolvedTrackedChangeType} después del paso ${stepIndex + 1} del replace.`,
        },
        recoveryAttempted: progress.recoveryAttempted,
        recoverySucceeded: progress.recoverySucceeded,
      };
    }

    const remainingTrackedChange = findTrackedChangeByType(
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
        recoveryAttempted: progress.recoveryAttempted,
        recoverySucceeded: progress.recoverySucceeded,
      };
    }

    return {
      observation: reobserved.observation,
      executionReport: {
        attempted: semanticOrder.length,
        completed: progress.completed,
        remaining: semanticOrder.length - progress.completed,
        failureIndex: stepIndex + 1,
        error:
          "Word reexpuso tracked changes incompatibles con el lado restante del replace.",
      },
      recoveryAttempted: progress.recoveryAttempted,
      recoverySucceeded: progress.recoverySucceeded,
    };
  }

  /** Collapses duplicate replace proxies into one semantic Deleted/Added pair before execution. */
  private normalizeReplaceObservation(
    observation: ResolutionObservation,
  ): ResolutionObservation {
    const normalizedTrackedChanges = [
      findTrackedChangeByType(observation.trackedChanges, "Deleted"),
      findTrackedChangeByType(observation.trackedChanges, "Added"),
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
    trackedChangeType: ReplaceTrackedChangeSide,
    report: ResolutionExecutionReport,
    workflowAttemptId: string,
  ): void {
    if (!report.unverifiedMutation || report.error) {
      return;
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace-step=${trackedChangeType} mutation verification unavailable after ${this.action}; re-observing fresh Word state before returning success`,
      report.unverifiedMutation,
    );
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
}
