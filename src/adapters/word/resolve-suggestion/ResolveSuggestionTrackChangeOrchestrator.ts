import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import type { SuggestionActionResult } from "../../../domain/suggestion/SuggestionResolutionWorkflow.types";
import type { DocumentReviewStateInspector } from "./DocumentReviewStateInspector";
import type {
  ResolutionObservation,
  ResolveSuggestionAction,
  ResolveSuggestionOutcome,
} from "./ResolutionContext";
import type { ResolutionErrorSerializer } from "./ResolutionErrorParser";
import type { ResolutionObservabilityReporter } from "./ResolutionObservabilityAdapter";
import { ResolveSuggestionOperationalExecutor } from "./ResolveSuggestionOperationalExecutor";
import type { ResolveSuggestionResultFactory } from "./ResolveSuggestionResultFactory";
import type { SuggestionLocator } from "./SuggestionLocator";
import type { SuggestionResolutionCleanup } from "./SuggestionResolutionCleanup";
import type { SuggestionResolutionObserver } from "./SuggestionResolutionObserver";

/**
 * Orchestrates the tracked-change resolution workflow.
 *
 * The flow is intentionally linear and fail-closed: locate wrapper, validate
 * the replace contract, observe actionable wrapper-owned tracked changes,
 * mutate that collection, clean the colocated comment, and inspect the final
 * review state.
 */
export class ResolveSuggestionTrackChangeOrchestrator {
  private readonly executor: ResolveSuggestionOperationalExecutor;

  constructor(
    private readonly suggestion: Suggestion,
    action: ResolveSuggestionAction,
    private readonly locator: SuggestionLocator,
    private readonly cleanup: SuggestionResolutionCleanup,
    private readonly observer: SuggestionResolutionObserver,
    private readonly resultFactory: ResolveSuggestionResultFactory,
    private readonly stateInspector: DocumentReviewStateInspector,
    private readonly observabilityReporter: ResolutionObservabilityReporter,
    errorSerializer: ResolutionErrorSerializer,
  ) {
    this.executor = new ResolveSuggestionOperationalExecutor(
      action,
      observabilityReporter,
      errorSerializer,
    );
  }

  /** Runs the tracked-change workflow inside the active Word batch. */
  async execute(
    context: Word.RequestContext,
  ): Promise<ResolveSuggestionOutcome> {
    await this.observabilityReporter.emitPhase("locate", "started", {
      suggestionType: this.suggestion.type,
    });

    const { candidates, selectedCc, locateStatus } =
      await this.locator.locateResolutionArtifacts(context);
    const pendingBefore = await this.stateInspector.inspect(context);

    await this.observabilityReporter.emitPhase(
      "locate",
      selectedCc ? "succeeded" : "failed",
      {
        candidateCount: candidates.length,
        selectedCcFound: Boolean(selectedCc),
        locateStatus,
      },
    );

    if (
      locateStatus === "ambiguous-location" ||
      locateStatus === "identity-lost"
    ) {
      const result = await this.resultFactory.buildObservationFailureResult(
        context,
        locateStatus,
        pendingBefore,
      );
      return this.toOutcome(result, pendingBefore);
    }

    if (!selectedCc) {
      return {
        status: "cc-not-found",
        trackedChangesAffected: 0,
        commentDeleted: false,
        pendingBefore,
        pendingAfter: pendingBefore,
      };
    }

    if (!this.hasValidTrackChangeContract()) {
      return this.toOutcome(
        this.resultFactory.buildErrorResult(
          "Contrato invalido de track-change: anchor y suggestedText son obligatorios.",
          pendingBefore,
          "observe-before",
        ),
        pendingBefore,
      );
    }

    const observationResult = await this.observeBeforeMutation(
      context,
      candidates,
      selectedCc,
      pendingBefore,
    );
    if (this.isWorkflowOutcome(observationResult)) {
      return observationResult;
    }

    const observation = observationResult;
    const executionReport = await this.executor.execute(context, observation);
    if (executionReport.error) {
      return this.toOutcome(
        this.resultFactory.buildErrorResult(
          executionReport.error,
          await this.stateInspector.inspect(context),
          "execute",
          executionReport,
        ),
        pendingBefore,
      );
    }

    const commentDeleted =
      await this.cleanup.deleteLocatedStylisticCommentAfterResolution(
        context,
        observation.selectedComment,
      );
    const pendingAfter =
      await this.stateInspector.inspectAfterResolution(context);

    await this.observabilityReporter.emitPhase("cleanup-comment", "succeeded", {
      commentDeleted,
    });
    await this.observabilityReporter.emitPhase("inspect-after", "succeeded", {
      pendingArtifacts: pendingAfter.pendingStylisticArtifacts,
    });

    return {
      status: this.resultFactory.toResolutionStatus(),
      trackedChangesAffected: executionReport.completed,
      commentDeleted,
      pendingBefore,
      pendingAfter,
      executionReport,
    };
  }

  /** Observes the wrapper state and degrades non-terminal host evidence before mutation. */
  private async observeBeforeMutation(
    context: Word.RequestContext,
    candidates: Word.ContentControl[],
    selectedCc: Word.ContentControl,
    pendingBefore: ResolveSuggestionOutcome["pendingBefore"],
  ): Promise<ResolutionObservation | ResolveSuggestionOutcome> {
    await this.observabilityReporter.emitPhase("observe-before", "started", {
      suggestionType: this.suggestion.type,
    });

    const observation = await this.observer.observeResolutionCandidates(
      context,
      candidates,
      selectedCc,
    );

    if (observation.observationStatus === "confirmed-pending") {
      await this.observabilityReporter.emitPhase(
        "observe-before",
        "succeeded",
        this.observabilityReporter.mergeMetadata(
          { trackedChangesObserved: observation.trackedChanges.length },
          observation.debugMetadata,
        ),
      );
      return observation;
    }

    const failureStatus = this.normalizeObservationFailure(observation);
    await this.observabilityReporter.emitPhase(
      "observe-before",
      failureStatus === "unobservable" ? "warning" : "failed",
      this.observabilityReporter.mergeMetadata(
        { reason: failureStatus },
        observation.debugMetadata,
      ),
    );

    const result = await this.resultFactory.buildObservationFailureResult(
      context,
      failureStatus,
      pendingBefore,
    );
    return this.toOutcome(result, pendingBefore);
  }

  /** Normalizes non-terminal observation variants to the public fail-closed statuses. */
  private normalizeObservationFailure(
    observation: ResolutionObservation,
  ): "identity-lost" | "ambiguous-location" | "mixed-group" | "unobservable" {
    if (
      observation.observationStatus === "identity-lost" ||
      observation.observationStatus === "ambiguous-location" ||
      observation.observationStatus === "mixed-group" ||
      observation.observationStatus === "unobservable"
    ) {
      return observation.observationStatus;
    }

    return "unobservable";
  }

  /** Returns true only when the suggestion satisfies the tracked-change replace contract. */
  private hasValidTrackChangeContract(): boolean {
    return (
      this.suggestion.type === "track-change" &&
      this.suggestion.anchor.trim().length > 0 &&
      (this.suggestion.suggestedText?.trim().length ?? 0) > 0
    );
  }

  /** Keeps the command-facing internal outcome shape stable. */
  private toOutcome(
    result: SuggestionActionResult,
    pendingBefore: ResolveSuggestionOutcome["pendingBefore"],
  ): ResolveSuggestionOutcome {
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

  /** Distinguishes terminal workflow outcomes from pre-execution observations. */
  private isWorkflowOutcome(
    value: ResolutionObservation | ResolveSuggestionOutcome,
  ): value is ResolveSuggestionOutcome {
    return "pendingBefore" in value;
  }
}
