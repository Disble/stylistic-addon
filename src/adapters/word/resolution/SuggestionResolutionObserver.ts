import type {
  ReplaceSuggestionIdentity,
  Suggestion,
  SuggestionObservationStatus,
} from "../../../domain/types";
import {
  getDeletedSideLocator,
  getOperationalAnchorLocator,
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
  ResolutionObservationDebugMetadata,
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

type ReplaceTrackedChangeSources = {
  ccTrackedChanges: Word.TrackedChange[];
  ccRangeTrackedChanges: Word.TrackedChange[];
  bodyRelatedTrackedChanges: Word.TrackedChange[];
  deletedSideTrackedChanges: Word.TrackedChange[];
  operationalAnchorTrackedChanges: Word.TrackedChange[];
  commentTrackedChanges: Word.TrackedChange[];
};

type LoadedReplaceObservationSources = {
  sources: ReplaceTrackedChangeSources;
  deletedSideText: string | null;
  operationalAnchorFound: boolean;
  baseDebugMetadata: Pick<
    ResolutionObservationDebugMetadata,
    | "ccTrackedChangesCount"
    | "ccRangeTrackedChangesCount"
    | "bodyTrackedChangesCount"
    | "bodyRelatedTrackedChangesCount"
    | "deletedSideTrackedChangesCount"
    | "deletedSideLocatorFound"
    | "operationalAnchorTrackedChangesCount"
    | "operationalAnchorFound"
    | "commentTrackedChangesCount"
  >;
};

/** Collects and classifies host evidence for one resolution workflow. */
export class SuggestionResolutionObserver {
  constructor(
    private readonly suggestion: Suggestion,
    private readonly locator: SuggestionLocator,
    private readonly textLocator: TextLocator,
  ) {}

  /** Builds one tracked-change diagnostic entry with its selected evidence source when known. */
  private describeTrackedChange(
    trackedChange: Word.TrackedChange,
    sources?: ReplaceTrackedChangeSources,
  ): {
    id: string;
    type: string;
    source?: string;
  } {
    return {
      id: String((trackedChange as { id?: string | number }).id ?? "no-id"),
      type: trackedChange.type ?? "unknown",
      ...(sources
        ? {
            source: this.identifyTrackedChangeSource(trackedChange, sources),
          }
        : {}),
    };
  }

  /** Builds a compact tracked-change list so one observation pass can be reconstructed later. */
  private describeTrackedChanges(
    trackedChanges: Word.TrackedChange[],
    sources?: ReplaceTrackedChangeSources,
  ): Array<{
    id: string;
    type: string;
    source?: string;
  }> {
    return trackedChanges.map((trackedChange) =>
      this.describeTrackedChange(trackedChange, sources),
    );
  }

  /** Logs every replace evidence source exactly once before the selector starts choosing between them. */
  private logReplaceSourceDiagnostics(
    cc: Word.ContentControl,
    loadedSources: LoadedReplaceObservationSources,
  ): void {
    console.log(
      `🧾 [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" replace-source-detail cc="${cc.tag}"`,
      {
        deletedSideText: loadedSources.deletedSideText,
        operationalAnchorFound: loadedSources.operationalAnchorFound,
        baseDebugMetadata: loadedSources.baseDebugMetadata,
        sources: {
          ccTrackedChanges: this.describeTrackedChanges(
            loadedSources.sources.ccTrackedChanges,
            loadedSources.sources,
          ),
          ccRangeTrackedChanges: this.describeTrackedChanges(
            loadedSources.sources.ccRangeTrackedChanges,
            loadedSources.sources,
          ),
          bodyRelatedTrackedChanges: this.describeTrackedChanges(
            loadedSources.sources.bodyRelatedTrackedChanges,
            loadedSources.sources,
          ),
          deletedSideTrackedChanges: this.describeTrackedChanges(
            loadedSources.sources.deletedSideTrackedChanges,
            loadedSources.sources,
          ),
          operationalAnchorTrackedChanges: this.describeTrackedChanges(
            loadedSources.sources.operationalAnchorTrackedChanges,
            loadedSources.sources,
          ),
          commentTrackedChanges: this.describeTrackedChanges(
            loadedSources.sources.commentTrackedChanges,
            loadedSources.sources,
          ),
        },
      },
    );
  }

  /** Logs how each replace-evidence combination scored before the selector returns one. */
  private logReplaceCandidateEvaluation(
    collections: Array<{
      label: string;
      trackedChanges: Word.TrackedChange[];
    }>,
    sources: ReplaceTrackedChangeSources,
  ): void {
    console.log(
      `🧭 [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" replace-candidate-evaluation`,
      {
        evaluations: collections.map(({ label, trackedChanges }) => ({
          label,
          count: trackedChanges.length,
          completePair:
            this.hasCompleteReplaceTrackedChangePair(trackedChanges),
          duplicateSide: this.hasDuplicateReplaceSide(trackedChanges),
          trackedChanges: this.describeTrackedChanges(trackedChanges, sources),
        })),
      },
    );
  }

  /** Returns `true` when a suggestion is semantically a replace operation. */
  private isReplaceSuggestion(): boolean {
    return (
      this.suggestion.type === "track-change" &&
      this.suggestion.anchor.length > 0 &&
      (this.suggestion.suggestedText?.length ?? 0) > 0
    );
  }

  /** Resolves one stable source label for a tracked change so post-step logs stay actionable. */
  private identifyTrackedChangeSource(
    trackedChange: Word.TrackedChange | null,
    sources: ReplaceTrackedChangeSources,
  ): string {
    if (!trackedChange) {
      return "none";
    }

    const trackedChangeId = String(
      (trackedChange as { id?: string | number }).id ?? "",
    );
    const belongsToSource = (collection: Word.TrackedChange[]): boolean =>
      collection.some((candidate) => {
        if (candidate === trackedChange) {
          return true;
        }

        const candidateId = String(
          (candidate as { id?: string | number }).id ?? "",
        );
        return trackedChangeId.length > 0 && candidateId === trackedChangeId;
      });

    if (belongsToSource(sources.deletedSideTrackedChanges)) {
      return "deletedSide";
    }

    if (belongsToSource(sources.operationalAnchorTrackedChanges)) {
      return "operationalAnchor";
    }

    if (belongsToSource(sources.ccTrackedChanges)) {
      return "cc";
    }

    if (belongsToSource(sources.ccRangeTrackedChanges)) {
      return "ccRange";
    }

    if (belongsToSource(sources.commentTrackedChanges)) {
      return "comment";
    }

    if (belongsToSource(sources.bodyRelatedTrackedChanges)) {
      return "bodyRelated";
    }

    return "unknown";
  }

  /**
   * Returns `true` only when replace evidence exposes BOTH semantic sides.
   *
   * A native Word replace tracked change is atomic from the user's perspective:
   * one side deletes the original text and the other inserts the replacement.
   * If Word currently exposes only one side, the suggestion is still pending and
   * the workflow must not certify a terminal accept/reject result yet.
   */
  private hasCompleteReplaceTrackedChangePair(
    trackedChanges: Word.TrackedChange[],
  ): boolean {
    let hasAdded = false;
    let hasDeleted = false;

    for (const trackedChange of trackedChanges) {
      if (trackedChange.type === "Added") {
        hasAdded = true;
      }

      if (trackedChange.type === "Deleted") {
        hasDeleted = true;
      }

      if (hasAdded && hasDeleted) {
        return true;
      }
    }

    return false;
  }

  /** Returns true when one replace side is duplicated inside the same evidence surface. */
  private hasDuplicateReplaceSide(
    trackedChanges: Word.TrackedChange[],
  ): boolean {
    let hasAdded = false;
    let hasDeleted = false;

    for (const trackedChange of trackedChanges) {
      if (trackedChange.type === "Added") {
        if (hasAdded) {
          return true;
        }

        hasAdded = true;
      }

      if (trackedChange.type === "Deleted") {
        if (hasDeleted) {
          return true;
        }

        hasDeleted = true;
      }
    }

    return false;
  }

  /**
   * Classifies which body tracked-changes are spatially related to the CC range.
   *
   * Word desktop sometimes throws `InvalidRibbonDefinition` (a misnamed
   * Office.js error code) from `compareLocationWith` when one of the body
   * tracked-change proxies references an invalidated range — typically
   * orphaned tracked-changes left by a prior failed reject. We isolate the
   * comparison sync so a single bad proxy cannot abort the entire observation;
   * we still have the ccTrackedChanges and ccRangeTrackedChanges evidence to
   * fall back on.
   */
  private async classifyBodyRelatedTrackedChanges(
    context: Word.RequestContext,
    candidates: Word.TrackedChange[],
    comparisons: OfficeExtension.ClientResult<string>[],
  ): Promise<Word.TrackedChange[]> {
    if (candidates.length === 0) {
      return [];
    }

    try {
      await context.sync();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" body tracked-change compareLocationWith sync failed (treating ${candidates.length} body candidates as not-related): ${message}`,
      );
      return [];
    }

    const related: Word.TrackedChange[] = [];
    for (let index = 0; index < candidates.length; index += 1) {
      let relationValue: string | undefined;
      try {
        relationValue = comparisons[index].value;
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        console.warn(
          `⚠️ [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" reading compareLocationWith result[${index}] failed: ${message}`,
        );
        continue;
      }
      if (RESOLUTION_RELATED_RELATIONS.has(relationValue)) {
        related.push(candidates[index]);
      }
    }
    return related;
  }

  /** Resolves all tracked changes semantically tied to a suggestion CC. */
  private async collectTrackedChangesForContentControl(
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ): Promise<{
    trackedChanges: Word.TrackedChange[];
    ccTrackedChanges: Word.TrackedChange[];
    ccRangeTrackedChanges: Word.TrackedChange[];
    bodyRelatedTrackedChanges: Word.TrackedChange[];
    debugMetadata: Pick<
      ResolutionObservationDebugMetadata,
      | "ccTrackedChangesCount"
      | "ccRangeTrackedChangesCount"
      | "bodyTrackedChangesCount"
      | "bodyRelatedTrackedChangesCount"
    >;
  }> {
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

    const bodyRelatedTrackedChanges =
      await this.classifyBodyRelatedTrackedChanges(
        context,
        candidateBodyTrackedChanges,
        comparisons,
      );
    for (const tc of bodyRelatedTrackedChanges) {
      addTrackedChange(tc);
    }

    const trackedChanges = [
      ...Array.from(trackedChangesById.values()),
      ...trackedChangesWithoutId,
    ];

    return {
      trackedChanges,
      ccTrackedChanges: ccTrackedChanges.items,
      ccRangeTrackedChanges: rangeTrackedChanges.items,
      bodyRelatedTrackedChanges,
      debugMetadata: {
        ccTrackedChangesCount: ccTrackedChanges.items.length,
        ccRangeTrackedChangesCount: rangeTrackedChanges.items.length,
        bodyTrackedChangesCount: bodyTrackedChanges.items.length,
        bodyRelatedTrackedChangesCount: bodyRelatedTrackedChanges.length,
      },
    };
  }

  /** Builds the fallback debug metadata for identity-lost and legacy observations. */
  private buildEmptyReplaceDebugMetadata(
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    observationStatus: SuggestionObservationStatus,
    identityVersion?: string,
  ): ResolutionObservationDebugMetadata {
    return {
      selectedCcTag: cc.tag,
      selectedCcTitleKind: (cc.title ?? "").startsWith("stylistic-meta-v2:")
        ? "compound-v2"
        : "legacy-or-empty",
      selectedCommentFound: Boolean(colocatedComment),
      trackedChangesObserved: 0,
      trackedChangeTypes: "",
      observationStatus,
      ...(identityVersion ? { identityVersion } : {}),
    };
  }

  /** Builds one identity-lost replace observation for structurally invalid compound-v2 metadata. */
  private buildIdentityLostReplaceObservation(
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    parsedIdentity: ReplaceSuggestionIdentity | null,
  ): ReplaceObservationContext {
    return {
      identity: parsedIdentity ?? undefined,
      trackedChanges: [],
      observationStatus: "identity-lost",
      debugMetadata: this.buildEmptyReplaceDebugMetadata(
        cc,
        colocatedComment,
        "identity-lost",
        parsedIdentity?.version,
      ),
    };
  }

  /** Loads every replace evidence source once so full and side-specific observation can share it. */
  private async loadReplaceObservationSources(
    context: Word.RequestContext,
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    parsedIdentity: ReplaceSuggestionIdentity,
  ): Promise<LoadedReplaceObservationSources> {
    const contentControlObservation =
      await this.collectTrackedChangesForContentControl(context, cc);

    const deletedSideText = getDeletedSideLocator(
      parsedIdentity,
      this.suggestion,
    );
    const deletedSideTrackedChanges = await this.loadDeletedSideTrackedChanges(
      context,
      deletedSideText,
    );

    const operationalAnchorRange = await this.resolveOperationalAnchorRange(
      context,
      parsedIdentity,
    );
    const operationalAnchorTrackedChanges = await this.loadRangeTrackedChanges(
      context,
      operationalAnchorRange,
    );

    const commentTrackedChanges = await this.loadCommentTrackedChanges(
      context,
      colocatedComment,
    );

    return {
      sources: {
        ccTrackedChanges: contentControlObservation.ccTrackedChanges,
        ccRangeTrackedChanges: contentControlObservation.ccRangeTrackedChanges,
        bodyRelatedTrackedChanges:
          contentControlObservation.bodyRelatedTrackedChanges,
        deletedSideTrackedChanges,
        operationalAnchorTrackedChanges,
        commentTrackedChanges,
      },
      deletedSideText,
      operationalAnchorFound: Boolean(operationalAnchorRange),
      baseDebugMetadata: {
        ccTrackedChangesCount:
          contentControlObservation.debugMetadata.ccTrackedChangesCount,
        ccRangeTrackedChangesCount:
          contentControlObservation.debugMetadata.ccRangeTrackedChangesCount,
        bodyTrackedChangesCount:
          contentControlObservation.debugMetadata.bodyTrackedChangesCount,
        bodyRelatedTrackedChangesCount:
          contentControlObservation.debugMetadata
            .bodyRelatedTrackedChangesCount,
        deletedSideTrackedChangesCount: deletedSideTrackedChanges.length,
        deletedSideLocatorFound: deletedSideText !== null,
        operationalAnchorTrackedChangesCount:
          operationalAnchorTrackedChanges.length,
        operationalAnchorFound: Boolean(operationalAnchorRange),
        commentTrackedChangesCount: commentTrackedChanges.length,
      },
    };
  }

  /** Loads tracked changes from the explicit deleted-side locator when present. */
  private async loadDeletedSideTrackedChanges(
    context: Word.RequestContext,
    deletedSideText: string | null,
  ): Promise<Word.TrackedChange[]> {
    if (!deletedSideText) {
      return [];
    }

    const deletedSideRange = await this.textLocator.locate({
      context,
      container: context.document.body as unknown as WordSearchContainer,
      searchText: deletedSideText,
    });

    return this.loadRangeTrackedChanges(context, deletedSideRange);
  }

  /** Loads tracked changes for one arbitrary range if the range exists. */
  private async loadRangeTrackedChanges(
    context: Word.RequestContext,
    range: Word.Range | null,
  ): Promise<Word.TrackedChange[]> {
    if (!range) {
      return [];
    }

    const trackedChanges = range.getTrackedChanges();
    trackedChanges.load({ select: "type,id" });
    await context.sync();
    return trackedChanges.items;
  }

  /** Loads tracked changes from the colocated Stylistic comment when present. */
  private async loadCommentTrackedChanges(
    context: Word.RequestContext,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<Word.TrackedChange[]> {
    if (!colocatedComment) {
      return [];
    }

    const commentRangeTrackedChanges =
      colocatedComment.range.getTrackedChanges();
    commentRangeTrackedChanges.load({ select: "type,id" });
    await context.sync();
    return commentRangeTrackedChanges.items;
  }

  /** Builds debug metadata for a fully observed replace pair. */
  private buildReplacePairDebugMetadata(
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    parsedIdentity: ReplaceSuggestionIdentity,
    trackedChanges: Word.TrackedChange[],
    sources: ReplaceTrackedChangeSources,
    baseDebugMetadata: LoadedReplaceObservationSources["baseDebugMetadata"],
  ): ResolutionObservationDebugMetadata {
    const observationStatus = this.hasCompleteReplaceTrackedChangePair(
      trackedChanges,
    )
      ? "confirmed-pending"
      : "unobservable";

    return {
      selectedCcTag: cc.tag,
      selectedCcTitleKind: "compound-v2",
      selectedCommentFound: Boolean(colocatedComment),
      trackedChangesObserved: trackedChanges.length,
      trackedChangeTypes: trackedChanges
        .map((trackedChange) => trackedChange.type ?? "unknown")
        .join(","),
      selectedDeletedSource: this.identifyTrackedChangeSource(
        trackedChanges.find(
          (trackedChange) => trackedChange.type === "Deleted",
        ) ?? null,
        sources,
      ),
      selectedAddedSource: this.identifyTrackedChangeSource(
        trackedChanges.find(
          (trackedChange) => trackedChange.type === "Added",
        ) ?? null,
        sources,
      ),
      observationStatus,
      identityVersion: parsedIdentity.version,
      ...baseDebugMetadata,
    };
  }

  /** Builds debug metadata for a side-specific replace observation. */
  private buildReplaceSemanticSideDebugMetadata(
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    parsedIdentity: ReplaceSuggestionIdentity,
    trackedChanges: Word.TrackedChange[],
    sources: ReplaceTrackedChangeSources,
    baseDebugMetadata: LoadedReplaceObservationSources["baseDebugMetadata"],
  ): ResolutionObservationDebugMetadata {
    const observationStatus =
      trackedChanges.length > 0 ? "confirmed-pending" : "unobservable";

    return {
      selectedCcTag: cc.tag,
      selectedCcTitleKind: "compound-v2",
      selectedCommentFound: Boolean(colocatedComment),
      trackedChangesObserved: trackedChanges.length,
      trackedChangeTypes: trackedChanges
        .map((trackedChange) => trackedChange.type ?? "unknown")
        .join(","),
      selectedSemanticSideSource: this.identifyTrackedChangeSource(
        trackedChanges[0] ?? null,
        sources,
      ),
      observationStatus,
      identityVersion: parsedIdentity.version,
      ...baseDebugMetadata,
    };
  }

  /** Returns the first evidence combination that yields a complete replace pair. */
  private resolveTrackedChangesForReplace(sources: {
    ccTrackedChanges: Word.TrackedChange[];
    ccRangeTrackedChanges: Word.TrackedChange[];
    bodyRelatedTrackedChanges: Word.TrackedChange[];
    deletedSideTrackedChanges: Word.TrackedChange[];
    operationalAnchorTrackedChanges: Word.TrackedChange[];
    commentTrackedChanges: Word.TrackedChange[];
  }): Word.TrackedChange[] {
    // Priority order is document-level first: `ccRange.getTrackedChanges()` and
    // the proximity-filtered `bodyRelated` stream operate on the real document
    // range that the user sees, so accepting their proxies actually mutates
    // the document. `cc.getTrackedChanges()` has been demoted to a later
    // fallback because in real Word it can expose CC-internal proxies that
    // silently no-op when accepted, leaving the replace pair unresolved while
    // `sync()` still succeeds. When the document-scoped sources cannot
    // complete the replace pair on their own, we combine the text-anchored
    // `deletedSide` locator with the remaining cc/ccRange evidence to avoid
    // regressing prior scenarios that relied on the text-anchored deletion.
    const prioritizedCollections: Array<{
      label: string;
      trackedChanges: Word.TrackedChange[];
    }> = [
      {
        label: "ccRange+bodyRelated",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccRangeTrackedChanges,
          sources.bodyRelatedTrackedChanges,
        ),
      },
      {
        label: "ccRange+deletedSide",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccRangeTrackedChanges,
          sources.deletedSideTrackedChanges,
        ),
      },
      {
        label: "cc+ccRange+deletedSide",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccTrackedChanges,
          sources.ccRangeTrackedChanges,
          sources.deletedSideTrackedChanges,
        ),
      },
      {
        label: "ccRange+bodyRelated+deletedSide",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccRangeTrackedChanges,
          sources.bodyRelatedTrackedChanges,
          sources.deletedSideTrackedChanges,
        ),
      },
      {
        label: "ccRange+bodyRelated+operationalAnchor",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccRangeTrackedChanges,
          sources.bodyRelatedTrackedChanges,
          sources.operationalAnchorTrackedChanges,
        ),
      },
      {
        label: "ccRange+bodyRelated+comment",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccRangeTrackedChanges,
          sources.bodyRelatedTrackedChanges,
          sources.commentTrackedChanges,
        ),
      },
      {
        label: "ccRange+operationalAnchor",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccRangeTrackedChanges,
          sources.operationalAnchorTrackedChanges,
        ),
      },
      {
        label: "ccRange+comment",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccRangeTrackedChanges,
          sources.commentTrackedChanges,
        ),
      },
      {
        label: "cc+ccRange+bodyRelated",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccTrackedChanges,
          sources.ccRangeTrackedChanges,
          sources.bodyRelatedTrackedChanges,
        ),
      },
      {
        label: "cc+ccRange+operationalAnchor",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccTrackedChanges,
          sources.ccRangeTrackedChanges,
          sources.operationalAnchorTrackedChanges,
        ),
      },
      {
        label: "cc+ccRange+comment",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccTrackedChanges,
          sources.ccRangeTrackedChanges,
          sources.commentTrackedChanges,
        ),
      },
      {
        label: "cc+ccRange+bodyRelated+deletedSide+operationalAnchor+comment",
        trackedChanges: this.mergeTrackedChanges(
          sources.ccTrackedChanges,
          sources.ccRangeTrackedChanges,
          sources.bodyRelatedTrackedChanges,
          sources.deletedSideTrackedChanges,
          sources.operationalAnchorTrackedChanges,
          sources.commentTrackedChanges,
        ),
      },
    ];
    this.logReplaceCandidateEvaluation(prioritizedCollections, sources);

    for (const { trackedChanges } of prioritizedCollections) {
      if (this.hasCompleteReplaceTrackedChangePair(trackedChanges)) {
        return trackedChanges;
      }
    }

    const lastPrioritizedCollection = [...prioritizedCollections].pop();

    return this.normalizeReplaceTrackedChanges(
      lastPrioritizedCollection?.trackedChanges ?? [],
    );
  }

  /** Relocates the operational anchor range persisted in compound-v2 metadata. */
  private async resolveOperationalAnchorRange(
    context: Word.RequestContext,
    identity: ReplaceSuggestionIdentity,
  ): Promise<Word.Range | null> {
    const anchorText =
      getOperationalAnchorLocator(identity, this.suggestion) ?? "";
    if (anchorText.length === 0) {
      return null;
    }

    return this.textLocator.locate({
      context,
      container: context.document.body as unknown as WordSearchContainer,
      searchText: anchorText,
    });
  }

  /** Merges tracked-change evidence from multiple sources without duplicating logical changes. */
  private mergeTrackedChanges(
    ...collections: Array<Word.TrackedChange[]>
  ): Word.TrackedChange[] {
    const trackedChangesById = new Map<string, Word.TrackedChange>();
    const trackedChangesWithoutId: Word.TrackedChange[] = [];

    for (const collection of collections) {
      for (const trackedChange of collection) {
        const id = String((trackedChange as { id?: string | number }).id ?? "");

        if (id.length > 0) {
          trackedChangesById.set(id, trackedChange);
          continue;
        }

        if (!trackedChangesWithoutId.includes(trackedChange)) {
          trackedChangesWithoutId.push(trackedChange);
        }
      }
    }

    return [
      ...Array.from(trackedChangesById.values()),
      ...trackedChangesWithoutId,
    ];
  }

  /** Keeps only one tracked change per semantic replace side. */
  private normalizeReplaceTrackedChanges(
    trackedChanges: Word.TrackedChange[],
  ): Word.TrackedChange[] {
    let selectedDeleted: Word.TrackedChange | null = null;
    let selectedAdded: Word.TrackedChange | null = null;

    for (const trackedChange of trackedChanges) {
      if (trackedChange.type === "Deleted" && selectedDeleted === null) {
        selectedDeleted = trackedChange;
      }

      if (trackedChange.type === "Added" && selectedAdded === null) {
        selectedAdded = trackedChange;
      }

      if (selectedDeleted && selectedAdded) {
        break;
      }
    }

    return [selectedDeleted, selectedAdded].filter(
      (trackedChange): trackedChange is Word.TrackedChange =>
        trackedChange !== null,
    );
  }

  /** Keeps only the first tracked change for one semantic side. */
  private normalizeTrackedChangesForSemanticSide(
    trackedChanges: Word.TrackedChange[],
    trackedChangeType: "Added" | "Deleted",
  ): Word.TrackedChange[] {
    const selectedTrackedChange = trackedChanges.find(
      (trackedChange) => trackedChange.type === trackedChangeType,
    );

    return selectedTrackedChange ? [selectedTrackedChange] : [];
  }

  /** Chooses the narrowest, side-specific evidence source for one remaining replace side. */
  private resolveTrackedChangesForReplaceSemanticSide(
    trackedChangeType: "Added" | "Deleted",
    sources: {
      ccTrackedChanges: Word.TrackedChange[];
      ccRangeTrackedChanges: Word.TrackedChange[];
      bodyRelatedTrackedChanges: Word.TrackedChange[];
      deletedSideTrackedChanges: Word.TrackedChange[];
      operationalAnchorTrackedChanges: Word.TrackedChange[];
      commentTrackedChanges: Word.TrackedChange[];
    },
  ): Word.TrackedChange[] {
    const prioritizedCollections =
      trackedChangeType === "Added"
        ? [
            sources.ccRangeTrackedChanges,
            sources.bodyRelatedTrackedChanges,
            sources.ccTrackedChanges,
            sources.operationalAnchorTrackedChanges,
            sources.commentTrackedChanges,
            this.mergeTrackedChanges(
              sources.ccRangeTrackedChanges,
              sources.bodyRelatedTrackedChanges,
              sources.ccTrackedChanges,
              sources.operationalAnchorTrackedChanges,
              sources.commentTrackedChanges,
            ),
          ]
        : [
            sources.bodyRelatedTrackedChanges,
            sources.ccRangeTrackedChanges,
            sources.deletedSideTrackedChanges,
            sources.operationalAnchorTrackedChanges,
            sources.commentTrackedChanges,
            sources.ccTrackedChanges,
            this.mergeTrackedChanges(
              sources.bodyRelatedTrackedChanges,
              sources.ccRangeTrackedChanges,
              sources.deletedSideTrackedChanges,
              sources.operationalAnchorTrackedChanges,
              sources.commentTrackedChanges,
              sources.ccTrackedChanges,
            ),
          ];

    for (const trackedChanges of prioritizedCollections) {
      const normalizedTrackedChanges =
        this.normalizeTrackedChangesForSemanticSide(
          trackedChanges,
          trackedChangeType,
        );

      if (normalizedTrackedChanges.length > 0) {
        return normalizedTrackedChanges;
      }
    }

    return [];
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
    if (!(cc.title ?? "").startsWith("stylistic-meta-v2:")) {
      return {
        trackedChanges: [],
        observationStatus: "unobservable",
        debugMetadata: this.buildEmptyReplaceDebugMetadata(
          cc,
          colocatedComment,
          "unobservable",
        ),
      };
    }

    if (!isValidCompoundReplaceIdentity(parsedIdentity, this.suggestion)) {
      return this.buildIdentityLostReplaceObservation(
        cc,
        colocatedComment,
        parsedIdentity,
      );
    }

    const loadedSources = await this.loadReplaceObservationSources(
      context,
      cc,
      colocatedComment,
      parsedIdentity,
    );
    this.logReplaceSourceDiagnostics(cc, loadedSources);
    const trackedChanges = this.resolveTrackedChangesForReplace(
      loadedSources.sources,
    );
    const debugMetadata = this.buildReplacePairDebugMetadata(
      cc,
      colocatedComment,
      parsedIdentity,
      trackedChanges,
      loadedSources.sources,
      loadedSources.baseDebugMetadata,
    );
    console.log(
      `🔬 [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" replace-selection cc="${cc.tag}"`,
      {
        selectedTrackedChanges: this.describeTrackedChanges(
          trackedChanges,
          loadedSources.sources,
        ),
        debugMetadata,
      },
    );

    return {
      identity: parsedIdentity,
      trackedChanges,
      observationStatus: debugMetadata.observationStatus ?? "unobservable",
      debugMetadata,
    };
  }

  /** Observes only one remaining semantic side of a replace after partial progress. */
  private async observeReplaceSuggestionSemanticSide(
    context: Word.RequestContext,
    cc: Word.ContentControl,
    colocatedComment: ColocatedCommentContext | null,
    trackedChangeType: "Added" | "Deleted",
  ): Promise<ReplaceObservationContext> {
    cc.load("title,tag");
    await context.sync();

    const parsedIdentity = parseReplaceIdentityTitle(
      (cc as { title?: string }).title,
    );
    if (!(cc.title ?? "").startsWith("stylistic-meta-v2:")) {
      return {
        trackedChanges: [],
        observationStatus: "unobservable",
        debugMetadata: this.buildEmptyReplaceDebugMetadata(
          cc,
          colocatedComment,
          "unobservable",
        ),
      };
    }

    if (!isValidCompoundReplaceIdentity(parsedIdentity, this.suggestion)) {
      return this.buildIdentityLostReplaceObservation(
        cc,
        colocatedComment,
        parsedIdentity,
      );
    }

    const loadedSources = await this.loadReplaceObservationSources(
      context,
      cc,
      colocatedComment,
      parsedIdentity,
    );
    const trackedChanges = this.resolveTrackedChangesForReplaceSemanticSide(
      trackedChangeType,
      loadedSources.sources,
    );
    const debugMetadata = this.buildReplaceSemanticSideDebugMetadata(
      cc,
      colocatedComment,
      parsedIdentity,
      trackedChanges,
      loadedSources.sources,
      loadedSources.baseDebugMetadata,
    );
    console.log(
      `🔬 [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" replace-semantic-selection cc="${cc.tag}" side=${trackedChangeType}`,
      {
        selectedTrackedChanges: this.describeTrackedChanges(
          trackedChanges,
          loadedSources.sources,
        ),
        debugMetadata,
      },
    );

    return {
      identity: parsedIdentity,
      trackedChanges,
      observationStatus: debugMetadata.observationStatus ?? "unobservable",
      debugMetadata,
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
    debugMetadata?: ResolutionObservationDebugMetadata;
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
        debugMetadata: observation.debugMetadata,
      };
    }

    const contentControlObservation =
      await this.collectTrackedChangesForContentControl(context, candidate);
    const trackedChanges = contentControlObservation.trackedChanges;
    const observationStatus =
      trackedChanges.length > 0 ? "confirmed-pending" : "unobservable";
    return {
      trackedChanges,
      observationStatus,
      debugMetadata: {
        selectedCcTag: candidate.tag,
        selectedCcTitleKind: (candidate.title ?? "").startsWith(
          "stylistic-meta-v2:",
        )
          ? "compound-v2"
          : "legacy-or-empty",
        selectedCommentFound: Boolean(colocatedComment),
        trackedChangesObserved: trackedChanges.length,
        trackedChangeTypes: trackedChanges
          .map((trackedChange) => trackedChange.type ?? "unknown")
          .join(","),
        observationStatus,
        ccTrackedChangesCount:
          contentControlObservation.debugMetadata.ccTrackedChangesCount,
        ccRangeTrackedChangesCount:
          contentControlObservation.debugMetadata.ccRangeTrackedChangesCount,
        bodyTrackedChangesCount:
          contentControlObservation.debugMetadata.bodyTrackedChangesCount,
        bodyRelatedTrackedChangesCount:
          contentControlObservation.debugMetadata
            .bodyRelatedTrackedChangesCount,
      },
    };
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
      console.log(
        `🧪 [SuggestionResolutionObserver] suggestionId="${this.suggestion.id}" candidate-observation tag="${candidate.tag}"`,
        {
          observationStatus: candidateObservation.observationStatus,
          trackedChangesObserved: candidateObservation.trackedChanges.length,
          debugMetadata: candidateObservation.debugMetadata ?? null,
        },
      );

      observation.selectedCc = candidate;
      observation.selectedComment = colocatedComment;
      observation.trackedChanges = candidateObservation.trackedChanges;
      observation.observationStatus = candidateObservation.observationStatus;
      observation.debugMetadata = candidateObservation.debugMetadata;

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

  /** Re-observes only one remaining replace side after one semantic step already succeeded. */
  async observeResolutionCandidatesForSemanticSide(
    context: Word.RequestContext,
    rankedCandidates: Word.ContentControl[],
    initialCc: Word.ContentControl,
    trackedChangeType: "Added" | "Deleted",
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
      const candidateObservation = this.isReplaceSuggestion()
        ? await this.observeReplaceSuggestionSemanticSide(
            context,
            candidate,
            colocatedComment,
            trackedChangeType,
          )
        : await this.observeResolutionCandidate(
            context,
            candidate,
            colocatedComment,
          );

      observation.selectedCc = candidate;
      observation.selectedComment = colocatedComment;
      observation.trackedChanges = candidateObservation.trackedChanges;
      observation.observationStatus = candidateObservation.observationStatus;
      observation.debugMetadata = candidateObservation.debugMetadata;

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
