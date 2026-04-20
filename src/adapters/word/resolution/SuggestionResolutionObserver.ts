import type {
  ReplaceSuggestionIdentity,
  Suggestion,
  SuggestionObservationStatus,
} from "../../../domain/types";
import {
  isValidCompoundReplaceIdentity,
  parseReplaceIdentityTitle,
} from "../ReplaceIdentityParser";
import type {
  TextLocator,
  WordSearchContainer,
} from "../WordTextLocatorContext";
import type {
  ColocatedCommentContext,
  ReplaceObservationContext,
  ResolutionObservation,
} from "./ResolutionContext";
import type { SuggestionLocator } from "./SuggestionLocator";

const RESOLUTION_RELATED_RELATIONS = new Set([
  "Contains",
  "ContainedIn",
  "OverlapsBefore",
  "OverlapsAfter",
  "Equal",
  "AdjacentBefore",
  "AdjacentAfter",
]);

/** Collects and classifies host evidence for one resolution workflow. */
export class SuggestionResolutionObserver {
  constructor(
    private readonly suggestion: Suggestion,
    private readonly locator: SuggestionLocator,
    private readonly textLocator: TextLocator,
  ) {}

  /** Returns `true` when a suggestion is semantically a replace operation. */
  private isReplaceSuggestion(): boolean {
    return (
      this.suggestion.type === "track-change" &&
      this.suggestion.anchor.length > 0 &&
      (this.suggestion.suggestedText?.length ?? 0) > 0
    );
  }

  /** Resolves all tracked changes semantically tied to a suggestion CC. */
  private async collectTrackedChangesForContentControl(
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ): Promise<Word.TrackedChange[]> {
    const ccRange = cc.getRange();

    const ccTrackedChanges = cc.getTrackedChanges();
    ccTrackedChanges.load({ select: "type,id" });

    const rangeTrackedChanges = ccRange.getTrackedChanges();
    rangeTrackedChanges.load({ select: "type,id" });

    const bodyTrackedChanges = context.document.body.getTrackedChanges();
    bodyTrackedChanges.load({ select: "type,id" });

    await context.sync();

    const trackedChangesById = new Map<string, Word.TrackedChange>();
    const trackedChangesWithoutId: Word.TrackedChange[] = [];

    const addTrackedChange = (tc: Word.TrackedChange) => {
      const id = String((tc as { id?: string | number }).id ?? "");
      if (id.length > 0) {
        trackedChangesById.set(id, tc);
      } else if (!trackedChangesWithoutId.includes(tc)) {
        trackedChangesWithoutId.push(tc);
      }
    };

    for (const tc of ccTrackedChanges.items) {
      addTrackedChange(tc);
    }

    for (const tc of rangeTrackedChanges.items) {
      addTrackedChange(tc);
    }

    const candidateBodyTrackedChanges = bodyTrackedChanges.items.filter(
      (tc) => {
        const id = String((tc as { id?: string | number }).id ?? "");
        return id.length === 0 || !trackedChangesById.has(id);
      },
    );

    const candidateRanges = candidateBodyTrackedChanges.map((tc) =>
      tc.getRange(),
    );
    const comparisons = candidateRanges.map((range) =>
      range.compareLocationWith(ccRange),
    );

    if (comparisons.length > 0) {
      await context.sync();
    }

    for (
      let index = 0;
      index < candidateBodyTrackedChanges.length;
      index += 1
    ) {
      const tc = candidateBodyTrackedChanges[index];
      if (
        RESOLUTION_RELATED_RELATIONS.has(comparisons[index].value as string)
      ) {
        addTrackedChange(tc);
      }
    }

    return [
      ...Array.from(trackedChangesById.values()),
      ...trackedChangesWithoutId,
    ];
  }

  /** Relocates the operational anchor range persisted in compound-v2 metadata. */
  private async resolveOperationalAnchorRange(
    context: Word.RequestContext,
    identity: ReplaceSuggestionIdentity,
  ): Promise<Word.Range | null> {
    const anchorText = identity.anchorRef?.value?.trim() ?? "";
    if (anchorText.length === 0) {
      return null;
    }

    return this.textLocator.locate({
      context,
      container: context.document.body as unknown as WordSearchContainer,
      searchText: anchorText,
    });
  }

  /** Observes replace suggestion evidence through compound-v2 metadata only. */
  private async observeReplaceSuggestion(
    context: Word.RequestContext,
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<ReplaceObservationContext> {
    cc.load("title,tag");
    await context.sync();

    const parsedIdentity = parseReplaceIdentityTitle(
      (cc as { title?: string }).title,
    );

    if ((cc as { title?: string }).title?.startsWith("stylistic-meta-v2:")) {
      if (!isValidCompoundReplaceIdentity(parsedIdentity, this.suggestion)) {
        return {
          identity: parsedIdentity ?? undefined,
          trackedChanges: [],
          observationStatus: "identity-lost",
        };
      }

      const trackedChanges = await this.collectTrackedChangesForContentControl(
        context,
        cc,
      );

      if (trackedChanges.length > 0) {
        return {
          identity: parsedIdentity,
          trackedChanges,
          observationStatus: "confirmed-pending",
        };
      }

      const operationalAnchorRange = await this.resolveOperationalAnchorRange(
        context,
        parsedIdentity,
      );

      if (operationalAnchorRange) {
        const anchorTrackedChanges = operationalAnchorRange.getTrackedChanges();
        anchorTrackedChanges.load({ select: "type,id" });
        await context.sync();

        if (anchorTrackedChanges.items.length > 0) {
          return {
            identity: parsedIdentity,
            trackedChanges: anchorTrackedChanges.items,
            observationStatus: "confirmed-pending",
          };
        }
      }

      if (colocatedComment) {
        const commentTrackedChanges =
          colocatedComment.range.getTrackedChanges();
        commentTrackedChanges.load({ select: "type,id" });
        await context.sync();

        if (commentTrackedChanges.items.length > 0) {
          return {
            identity: parsedIdentity,
            trackedChanges: commentTrackedChanges.items,
            observationStatus: "confirmed-pending",
          };
        }
      }

      return {
        identity: parsedIdentity,
        trackedChanges: [],
        observationStatus: "unobservable",
      };
    }

    return {
      trackedChanges: [],
      observationStatus: "unobservable",
    };
  }

  /** Observes one resolution candidate and normalizes replace vs non-replace evidence. */
  private async observeResolutionCandidate(
    context: Word.RequestContext,
    candidate: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<{
    trackedChanges: Word.TrackedChange[];
    observationStatus: SuggestionObservationStatus;
  }> {
    if (this.isReplaceSuggestion()) {
      const observation = await this.observeReplaceSuggestion(
        context,
        candidate,
        colocatedComment,
      );
      return {
        trackedChanges: observation.trackedChanges,
        observationStatus: observation.observationStatus,
      };
    }

    const trackedChanges = await this.collectTrackedChangesForContentControl(
      context,
      candidate,
    );
    const observationStatus =
      trackedChanges.length > 0 ? "confirmed-pending" : "unobservable";
    return { trackedChanges, observationStatus };
  }

  /** Chooses the best observed candidate by scanning ranked CCs until evidence is conclusive. */
  async observeResolutionCandidates(
    context: Word.RequestContext,
    rankedCandidates: Word.ContentControl[],
    initialCc: Word.ContentControl,
  ): Promise<ResolutionObservation> {
    const observation: ResolutionObservation = {
      selectedCc: initialCc,
      selectedComment: null,
      trackedChanges: [],
      observationStatus: "unobservable",
    };

    for (const candidate of rankedCandidates) {
      const colocatedComment = await this.locator.findColocatedStylisticComment(
        context,
        candidate,
      );
      const candidateObservation = await this.observeResolutionCandidate(
        context,
        candidate,
        colocatedComment,
      );

      observation.selectedCc = candidate;
      observation.selectedComment = colocatedComment;
      observation.trackedChanges = candidateObservation.trackedChanges;
      observation.observationStatus = candidateObservation.observationStatus;

      if (
        candidateObservation.observationStatus === "identity-lost" ||
        (candidateObservation.observationStatus === "confirmed-pending" &&
          candidateObservation.trackedChanges.length > 0)
      ) {
        break;
      }
    }

    return observation;
  }
}
