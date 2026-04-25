/* global console */

import type { ResolutionExecutionReport } from "../../../domain/types";
import type {
  ReplaceResolutionStrategy,
  ReplaceTrackedChangeSide,
} from "./ReplaceResolutionStrategyContext";
import type {
  ReplaceTrackedChangeCandidate,
  ResolutionObservation,
} from "./ResolutionContext";
import {
  formatTrackedChangeTypesForLog,
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
};

type ReplaceProgressState = {
  completed: number;
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
    const semanticOrder = this.replaceResolutionStrategy.semanticOrder;
    let activeObservation = observation;
    let completed = 0;

    for (const [stepIndex, trackedChangeType] of semanticOrder.entries()) {
      const stepResult = await this.executeSemanticStep(
        context,
        activeObservation,
        trackedChangeType,
        workflowAttemptId,
      );

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
        };
      }

      if (stepIndex < semanticOrder.length - 1) {
        const reobservationResult = await this.reobserveRemainingReplaceSide(
          context,
          activeObservation,
          semanticOrder,
          stepIndex,
          { completed },
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
    };
  }

  /** Executes one semantic side of a replace and fails closed when Word cannot certify it. */
  private async executeSemanticStep(
    context: Word.RequestContext,
    observation: ResolutionObservation,
    trackedChangeType: ReplaceTrackedChangeSide,
    workflowAttemptId: string,
  ): Promise<{
    observation: ResolutionObservation;
    completed: boolean;
    error?: string;
  }> {
    const initialCandidates = this.getSemanticCandidates(
      observation,
      trackedChangeType,
    );

    if (initialCandidates.length === 0) {
      return {
        observation,
        completed: false,
        error: `Word no reexpuso el tracked change ${trackedChangeType} requerido para resolver el replace.`,
      };
    }

    const initialAttempt = await this.executeSemanticCandidates(
      context,
      observation,
      trackedChangeType,
      initialCandidates,
      workflowAttemptId,
    );
    if (initialAttempt.completed) {
      return {
        observation,
        completed: true,
      };
    }

    this.logUnverifiedReplaceSemanticStep(
      trackedChangeType,
      initialAttempt.report,
      workflowAttemptId,
    );
    const initialErrorMessage = this.buildReplaceSemanticStepErrorMessage(
      trackedChangeType,
      initialAttempt.report,
    );

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace-step=${trackedChangeType} failed without recovery: ${initialErrorMessage}`,
    );

    return {
      observation,
      completed: false,
      error: initialErrorMessage,
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

  /** Returns every executable candidate exposed for one semantic replace side. */
  private getSemanticCandidates(
    observation: ResolutionObservation,
    trackedChangeType: ReplaceTrackedChangeSide,
  ): ReplaceTrackedChangeCandidate[] {
    return observation.semanticCandidates?.[trackedChangeType] ?? [];
  }

  /** Executes candidates in sequence until one is semantically verified or the list is exhausted. */
  private async executeSemanticCandidates(
    context: Word.RequestContext,
    observation: ResolutionObservation,
    trackedChangeType: ReplaceTrackedChangeSide,
    candidates: ReplaceTrackedChangeCandidate[],
    workflowAttemptId: string,
  ): Promise<{
    completed: boolean;
    report: ResolutionExecutionReport;
  }> {
    let lastReport: ResolutionExecutionReport | undefined;

    for (const candidate of candidates) {
      console.log(
        `🧪 [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace-step=${trackedChangeType} execute source=${candidate.source} status=${observation.observationStatus} trackedChanges=${observation.trackedChanges.length} types=${formatTrackedChangeTypesForLog(observation.trackedChanges)}`,
      );

      const report = await this.executor.apply(context, [
        candidate.trackedChange,
      ]);
      if (this.isExecutionReportSemanticallyVerified(report)) {
        return {
          completed: true,
          report,
        };
      }

      lastReport = report;
    }

    return {
      completed: false,
      report: lastReport ?? {
        attempted: 0,
        completed: 0,
        remaining: 0,
        error: `Word no expuso candidatos ejecutables para el tracked change ${trackedChangeType}.`,
      },
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
      };
    }

    console.log(
      `🔁 [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace re-observation step=${stepIndex + 1} side=${remainingTrackedChangeType} status=${reobserved.observation.observationStatus} trackedChanges=${reobserved.observation.trackedChanges.length} types=${formatTrackedChangeTypesForLog(reobserved.observation.trackedChanges)} deletedSources=${this.describeSemanticCandidateSources(reobserved.observation, "Deleted")} addedSources=${this.describeSemanticCandidateSources(reobserved.observation, "Added")}`,
    );

    if (
      this.getSemanticCandidates(
        reobserved.observation,
        resolvedTrackedChangeType,
      ).length > 0
    ) {
      return {
        observation: reobserved.observation,
        executionReport: {
          attempted: semanticOrder.length,
          completed: progress.completed,
          remaining: semanticOrder.length - progress.completed,
          failureIndex: stepIndex,
          error: `Word mantuvo pendiente el tracked change ${resolvedTrackedChangeType} después del paso ${stepIndex + 1} del replace.`,
        },
      };
    }

    if (
      this.getSemanticCandidates(
        reobserved.observation,
        remainingTrackedChangeType,
      ).length > 0
    ) {
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
    };
  }

  /** Formats semantic candidate sources directly from the observation instead of duplicating them in debug metadata. */
  private describeSemanticCandidateSources(
    observation: ResolutionObservation,
    trackedChangeType: ReplaceTrackedChangeSide,
  ): string {
    const sources = this.getSemanticCandidates(
      observation,
      trackedChangeType,
    ).map((candidate) => candidate.source);

    return sources.length > 0 ? sources.join(",") : "none";
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
  ): string {
    if (report.error) {
      return report.error;
    }

    const actionLabel = this.replaceResolutionStrategy.actionLabel;
    if (report.unverifiedMutation) {
      return `Word no pudo verificar si el ${actionLabel} del lado ${trackedChangeType} mutó el documento (${this.formatUnverifiedMutationForLog(report.unverifiedMutation)}).`;
    }

    return `Word ignoró el ${actionLabel} del lado ${trackedChangeType} (silent no-op detectado: el proxy del tracked change no mutó el documento).`;
  }

  /** Logs when Word mutated a replace side but the workflow now fails closed instead of retrying. */
  private logUnverifiedReplaceSemanticStep(
    trackedChangeType: ReplaceTrackedChangeSide,
    report: ResolutionExecutionReport,
    workflowAttemptId: string,
  ): void {
    if (!report.unverifiedMutation || report.error) {
      return;
    }

    console.warn(
      `⚠️ [ResolveSuggestionCommand] workflowAttemptId="${workflowAttemptId}" replace-step=${trackedChangeType} mutation verification unavailable after ${this.action}; failing closed without recovery`,
      report.unverifiedMutation,
    );
  }
}
