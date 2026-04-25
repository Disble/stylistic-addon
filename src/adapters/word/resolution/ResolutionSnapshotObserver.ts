import type {
  DocumentReviewState,
  ResolutionObservabilitySnapshotLabel,
  Suggestion,
} from "../../../domain/types";
import type { DocumentReviewStateInspector } from "./DocumentReviewStateInspector";
import type { ResolutionObservation } from "./ResolutionContext";
import type { ResolutionErrorSerializer } from "./ResolutionErrorParser";
import type { ResolutionObservabilityReporter } from "./ResolutionObservabilityAdapter";
import {
  describeTrackedChangesForLog,
  prioritizeFreshPreferredCandidate,
  resolveFreshPreferredCandidate,
} from "./ResolutionObservationContext";
import type { SuggestionLocator } from "./SuggestionLocator";
import type { SuggestionResolutionObserver } from "./SuggestionResolutionObserver";

/** Captures best-effort workflow snapshots without keeping observability assembly inside the command. */
export class ResolutionSnapshotObserver {
  constructor(
    private readonly suggestion: Suggestion,
    private readonly stateInspector: DocumentReviewStateInspector,
    private readonly locator: SuggestionLocator,
    private readonly observer: SuggestionResolutionObserver,
    private readonly observabilityReporter: ResolutionObservabilityReporter,
    private readonly errorSerializer: ResolutionErrorSerializer,
  ) {}

  /** Captures one fresh workflow snapshot around execute or cleanup and returns any fresh observation produced. */
  async capture(
    context: Word.RequestContext,
    label: ResolutionObservabilitySnapshotLabel,
    preferredCc?: Word.ContentControl,
    reviewState?: DocumentReviewState,
  ): Promise<ResolutionObservation | null> {
    try {
      const currentReviewState =
        reviewState ?? (await this.stateInspector.inspect(context));

      if (this.suggestion.type === "comment-only") {
        await this.observabilityReporter.captureSnapshot(label, {
          reviewState: currentReviewState,
          replaceSuggestion: false,
        });
        return null;
      }

      const relocated = await this.locator.locateResolutionArtifacts(context);
      if (!relocated.selectedCc) {
        await this.observabilityReporter.captureSnapshot(
          label,
          {
            reviewState: currentReviewState,
            replaceSuggestion: true,
          },
          {
            relocatedCandidateCount: relocated.candidates.length,
            relocatedSelectedCc: null,
          },
        );
        return null;
      }

      const resolvedPreferredCc = resolveFreshPreferredCandidate(
        relocated.candidates,
        preferredCc,
      );
      const preferredCandidates = prioritizeFreshPreferredCandidate(
        relocated.candidates,
        resolvedPreferredCc,
      );
      const observation = await this.observer.observeResolutionCandidates(
        context,
        preferredCandidates,
        resolvedPreferredCc ?? relocated.selectedCc,
      );

      await this.observabilityReporter.captureSnapshot(
        label,
        {
          reviewState: currentReviewState,
          replaceSuggestion: true,
          observationStatus: observation.observationStatus,
        },
        {
          relocatedCandidateCount: relocated.candidates.length,
          relocatedSelectedCc: relocated.selectedCc.tag,
          preferredCcResolved: resolvedPreferredCc?.tag ?? null,
          trackedChangesObserved: observation.trackedChanges.length,
          trackedChanges: describeTrackedChangesForLog(
            observation.trackedChanges,
          ),
        },
        {
          debugMetadata: observation.debugMetadata ?? null,
        },
      );
      return observation;
    } catch (error) {
      const serializedError = this.errorSerializer.serialize(error);
      await this.observabilityReporter.captureSnapshot(
        label,
        {
          replaceSuggestion: this.suggestion.type === "track-change",
          snapshotFailed: true,
          snapshotError: serializedError.message,
        },
        undefined,
        {
          error: serializedError,
        },
      );
      return null;
    }
  }
}
