import type {
  ReplaceSuggestionIdentity,
  Suggestion,
  SuggestionObservationStatus,
} from "../../../domain/types";
import {
  isValidOperationalReplaceIdentity,
  parseReplaceIdentityTitle,
} from "../ReplaceIdentityParser";
import type { TextLocator } from "../WordTextLocatorContext";
import { OperationalWrapperGroupResolver } from "./OperationalWrapperGroupResolver";
import type {
  ColocatedCommentContext,
  ReplaceObservationContext,
  ResolutionObservation,
  ResolutionObservationDebugMetadata,
  ResolutionTrackedChangeCollection,
} from "./ResolutionContext";
import type { SuggestionLocator } from "./SuggestionLocator";

type ReplaceTrackedChangeSide = "Added" | "Deleted";

/**
 * Observes resolution state strictly from operational wrapper scopes.
 *
 * It deliberately rejects the old strategy of reconstructing replace semantics
 * from partial host evidence (`cc`, `ccRange`, `comment`, `Deleted`, etc.).
 */
export class SuggestionResolutionObserver {
  private readonly groupResolver = new OperationalWrapperGroupResolver();

  constructor(
    private readonly suggestion: Suggestion,
    private readonly locator: SuggestionLocator,
    private readonly _textLocator: TextLocator,
  ) {}

  /** Observes whether the selected wrapper exposes one executable operational scope. */
  async observeResolutionCandidates(
    context: Word.RequestContext,
    _candidates: Word.ContentControl[],
    selectedCc: Word.ContentControl,
  ): Promise<ResolutionObservation> {
    const selectedComment = await this.locator.findColocatedStylisticComment(
      context,
      selectedCc,
    );
    const identity = parseReplaceIdentityTitle(selectedCc.title);

    if (!isValidOperationalReplaceIdentity(identity, this.suggestion)) {
      return this.buildAbortObservation(
        selectedCc,
        selectedComment,
        "identity-lost",
        identity,
      );
    }

    const group = await this.groupResolver.resolve(
      context,
      selectedCc,
      identity,
    );
    if (group.status === "ambiguous") {
      return this.buildAbortObservation(
        selectedCc,
        selectedComment,
        "ambiguous-location",
        identity,
        group,
      );
    }

    if (group.status === "contiguous") {
      return this.buildAbortObservation(
        selectedCc,
        selectedComment,
        "mixed-group",
        identity,
        group,
      );
    }

    const trackedChangesCollection = this.loadWrapperTrackedChanges(selectedCc);
    await context.sync();
    const trackedChanges = [...trackedChangesCollection.items];
    const observationStatus =
      trackedChanges.length > 0 ? "confirmed-pending" : "unobservable";

    return {
      selectedCc,
      selectedComment,
      trackedChanges,
      trackedChangesCollection,
      observationStatus,
      debugMetadata: this.buildDebugMetadata(
        selectedCc,
        selectedComment,
        observationStatus,
        identity,
        group,
        trackedChanges.length,
      ),
      group,
    };
  }

  /** Filters the wrapper scope to one semantic side for tests/snapshots that still need it. */
  async observeResolutionCandidatesForSemanticSide(
    context: Word.RequestContext,
    candidates: Word.ContentControl[],
    selectedCc: Word.ContentControl,
    trackedChangeType: ReplaceTrackedChangeSide,
  ): Promise<ResolutionObservation> {
    const observation = await this.observeResolutionCandidates(
      context,
      candidates,
      selectedCc,
    );

    if (observation.observationStatus !== "confirmed-pending") {
      return observation;
    }

    const trackedChanges = observation.trackedChanges.filter(
      (trackedChange) => trackedChange.type === trackedChangeType,
    );

    return {
      ...observation,
      trackedChanges,
      observationStatus:
        trackedChanges.length > 0 ? "confirmed-pending" : "unobservable",
      debugMetadata: {
        ...observation.debugMetadata,
        observationStatus:
          trackedChanges.length > 0 ? "confirmed-pending" : "unobservable",
      },
    };
  }

  /** Loads the wrapper range tracked-change collection that owns the operation. */
  private loadWrapperTrackedChanges(
    selectedCc: Word.ContentControl,
  ): ResolutionTrackedChangeCollection {
    const trackedChanges = selectedCc
      .getRange()
      .getTrackedChanges() as ResolutionTrackedChangeCollection;
    trackedChanges.load({ select: "type,id" });
    return trackedChanges;
  }

  /** Builds one fail-closed observation before any mutation happens. */
  private buildAbortObservation(
    selectedCc: Word.ContentControl,
    selectedComment: ColocatedCommentContext | null,
    observationStatus: Exclude<
      SuggestionObservationStatus,
      "confirmed-pending" | "confirmed-resolved" | "unobservable"
    >,
    identity: ReplaceSuggestionIdentity | null,
    group?: ReplaceObservationContext["group"],
  ): ResolutionObservation {
    return {
      selectedCc,
      selectedComment,
      trackedChanges: [],
      observationStatus,
      debugMetadata: this.buildDebugMetadata(
        selectedCc,
        selectedComment,
        observationStatus,
        identity,
        group,
        0,
      ),
      ...(group ? { group } : {}),
    };
  }

  /** Builds stable non-heuristic debug metadata for observability. */
  private buildDebugMetadata(
    selectedCc: Word.ContentControl,
    selectedComment: ColocatedCommentContext | null,
    observationStatus: SuggestionObservationStatus,
    identity: ReplaceSuggestionIdentity | null,
    group: ReplaceObservationContext["group"] | undefined,
    trackedChangesCount: number,
  ): ResolutionObservationDebugMetadata {
    return {
      selectedCcTag: selectedCc.tag,
      selectedCcTitleKind:
        identity?.version === "operational-wrapper-v1"
          ? "operational-wrapper-v1"
          : "invalid-or-missing",
      selectedCommentFound: Boolean(selectedComment),
      observationStatus,
      ...(identity ? { identityVersion: identity.version } : {}),
      ccRangeTrackedChangesCount: trackedChangesCount,
      ...(group
        ? {
            wrapperGroupId: group.groupId,
            wrapperGroupSize: group.members.length,
            wrapperGroupStatus: group.status,
          }
        : {}),
    };
  }
}
