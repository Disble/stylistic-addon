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

import {
  DocumentReviewStateMachine,
  type DocumentReviewUiState,
} from "../../domain/review/DocumentReviewStateMachine";
import type {
  DocumentReviewState,
  ReplaceSuggestionIdentity,
  Suggestion,
  SuggestionActionResult,
  SuggestionObservationStatus,
} from "../../domain/types";
import {
  STYLISTIC_IDENTITY_TITLE_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../infrastructure/config";
import { OVERLAPPING_RELATIONS } from "./cleanup/CommentCleanup";
import { isStylisticComment } from "./StylisticCommentBuilder";

// ---------------------------------------------------------------------------
// Exported identity utilities — shared with WordAdapter.getAppliedOriginalTexts
// ---------------------------------------------------------------------------

/**
 * Parses persisted v2 replace identity metadata from a CC title payload.
 * Returns `null` when the title does not carry a valid compound-v2 prefix.
 */
export function parseReplaceIdentityTitle(
  title: string | undefined,
): ReplaceSuggestionIdentity | null {
  const trimmed = title?.trim() ?? "";
  if (!trimmed.startsWith(STYLISTIC_IDENTITY_TITLE_PREFIX)) {
    return null;
  }

  const rawPayload = trimmed.slice(STYLISTIC_IDENTITY_TITLE_PREFIX.length);
  if (rawPayload.length === 0) {
    return null;
  }

  try {
    return JSON.parse(rawPayload) as ReplaceSuggestionIdentity;
  } catch {
    return null;
  }
}

/**
 * Validates the minimum compound-v2 replace identity contract required for
 * safe observation.
 */
export function isValidCompoundReplaceIdentity(
  identity: ReplaceSuggestionIdentity | null,
  suggestion: Suggestion,
): identity is ReplaceSuggestionIdentity {
  if (identity?.version !== "compound-v2") {
    return false;
  }

  return (
    identity.suggestionId === suggestion.id &&
    identity.insertedSideRef?.kind === "content-control" &&
    identity.insertedSideRef.role === "inserted-side" &&
    identity.insertedSideRef.value.length > 0 &&
    identity.deletedSideRef?.role === "deleted-side" &&
    identity.deletedSideRef.value.length > 0 &&
    identity.anchorRef?.role === "operational-anchor" &&
    identity.anchorRef.value.length > 0
  );
}

// ---------------------------------------------------------------------------
// Internal types
// ---------------------------------------------------------------------------

type ResolutionStatus =
  | "accepted"
  | "rejected"
  | "already-resolved"
  | "unobservable"
  | "identity-lost";

type ReplaceObservationContext = {
  identity?: ReplaceSuggestionIdentity;
  trackedChanges: Word.TrackedChange[];
  observationStatus: SuggestionObservationStatus;
};

type ColocatedCommentContext = {
  comment: Word.Comment;
  range: Word.Range;
};

type ResolutionObservation = {
  selectedCc: Word.ContentControl;
  selectedComment: ColocatedCommentContext | null;
  trackedChanges: Word.TrackedChange[];
  observationStatus: SuggestionObservationStatus;
};

type SearchContainer = {
  search(text: string, options: Record<string, boolean>): Word.RangeCollection;
  load(property: "text"): void;
  text: string;
};

type IndexedText = {
  text: string;
  indices: number[];
};

const RESOLUTION_RELATED_RELATIONS = new Set([
  ...OVERLAPPING_RELATIONS,
  "AdjacentBefore",
  "AdjacentAfter",
]);

// ---------------------------------------------------------------------------
// Pure utility functions
// ---------------------------------------------------------------------------

function removeWhitespaceWithIndices(text: string): IndexedText {
  const indices: number[] = [];
  let normalized = "";

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (/\s/.test(char)) {
      continue;
    }

    normalized += char;
    indices.push(index);
  }

  return { text: normalized, indices };
}

function findWhitespaceInsensitiveSlice(
  searchText: string,
  documentText: string,
): string | null {
  const normalizedSearch = removeWhitespaceWithIndices(searchText).text;
  if (normalizedSearch.length === 0) {
    return null;
  }

  const normalizedDocument = removeWhitespaceWithIndices(documentText);
  const matchIndex = normalizedDocument.text.indexOf(normalizedSearch);
  if (matchIndex === -1) {
    return null;
  }

  const start = normalizedDocument.indices[matchIndex];
  const end =
    normalizedDocument.indices[matchIndex + normalizedSearch.length - 1] + 1;
  return documentText.slice(start, end);
}

/** Returns `true` when a suggestion is semantically a replace operation. */
function isReplaceSuggestion(suggestion: Suggestion): boolean {
  return (
    suggestion.type === "track-change" &&
    suggestion.anchor.length > 0 &&
    (suggestion.suggestedText?.length ?? 0) > 0
  );
}

// ---------------------------------------------------------------------------
// Document review state helpers (shared logic extracted as functions)
// ---------------------------------------------------------------------------

/** Creates a normalized document-review snapshot. */
function buildDocumentReviewState(
  pendingStylisticArtifacts: number,
  trackChangesActive: boolean,
): DocumentReviewState {
  return {
    pendingStylisticArtifacts,
    hasPendingStylisticArtifacts: pendingStylisticArtifacts > 0,
    trackChangesActive,
  };
}

/** Derives the explicit document-review UI state from a document snapshot. */
function deriveDocumentState(
  reviewState: DocumentReviewState,
): DocumentReviewUiState {
  return DocumentReviewStateMachine.deriveState(reviewState);
}

/** Reads the authoritative document-derived review state in the current batch. */
async function inspectDocumentReviewState(
  context: Word.RequestContext,
): Promise<DocumentReviewState> {
  const allCCs = context.document.contentControls;
  allCCs.load("items/tag");
  context.document.load("changeTrackingMode");
  await context.sync();

  const pendingStylisticArtifacts = allCCs.items.filter((cc) =>
    cc.tag.startsWith(STYLISTIC_TAG_PREFIX),
  ).length;
  const trackChangesActive =
    context.document.changeTrackingMode !== Word.ChangeTrackingMode.off;

  return buildDocumentReviewState(
    pendingStylisticArtifacts,
    trackChangesActive,
  );
}

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
  constructor(
    private readonly suggestion: Suggestion,
    private readonly action: "accept" | "reject",
  ) {}

  /**
   * Executes the resolution command.
   * Never throws — catches all errors and returns a result object.
   */
  async execute(): Promise<SuggestionActionResult> {
    try {
      return await Word.run(async (context) => {
        console.log(
          `🎯 [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" type=${this.suggestion.type}`,
        );
        const pendingBefore = await inspectDocumentReviewState(context);

        const ccResult = this.findCCByTag(context);
        await context.sync();

        console.log(
          `🎯 [ResolveSuggestionCommand] getByTag returned ${ccResult.items.length} CC candidate(s) for suggestionId="${this.suggestion.id}"`,
        );

        const rankedCandidates = this.rankResolutionContentControls(
          ccResult.items,
        );
        const cc = this.selectResolutionContentControl(rankedCandidates);

        if (!cc) {
          console.warn(
            `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" failed: CC not found`,
          );
          return this.buildResolutionResult(
            "cc-not-found",
            0,
            false,
            pendingBefore,
            pendingBefore,
          );
        }

        console.log(
          `🎯 [ResolveSuggestionCommand] selected CC for suggestionId="${this.suggestion.id}": tag="${cc.tag}" hasV2Title=${Boolean(cc.title?.startsWith(STYLISTIC_IDENTITY_TITLE_PREFIX))}`,
        );

        if (this.suggestion.type === "comment-only") {
          const colocatedComment = await this.findColocatedStylisticComment(
            context,
            cc,
          );
          const commentDeleted = await this.deleteLocatedStylisticComment(
            context,
            colocatedComment,
          );
          return this.resolveCommentOnlySuggestion(
            context,
            cc,
            commentDeleted,
            pendingBefore,
          );
        }

        const observation = await this.observeResolutionCandidates(
          context,
          rankedCandidates,
          cc,
        );

        if (observation.observationStatus === "identity-lost") {
          console.warn(
            `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" ended in identity-lost`,
          );
          return this.buildObservationFailureResult(
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
          return this.buildObservationFailureResult(
            context,
            "unobservable",
            pendingBefore,
          );
        }

        this.applyTrackedChangeResolution(observation.trackedChanges);

        const commentDeleted =
          await this.deleteLocatedStylisticCommentAfterResolution(
            context,
            observation.selectedComment,
          );

        await this.cleanupResolvedSuggestionAnchor(
          context,
          observation.selectedCc,
        );

        const pendingAfter =
          await this.inspectDocumentReviewStateAfterResolution(
            context,
            pendingBefore,
          );

        return this.buildResolutionResult(
          this.toResolutionStatus(),
          observation.trackedChanges.length,
          commentDeleted,
          pendingBefore,
          pendingAfter,
        );
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [ResolveSuggestionCommand] threw for suggestionId="${this.suggestion.id}": ${message}`,
      );
      const pendingAfter = await Word.run((ctx) =>
        inspectDocumentReviewState(ctx),
      ).catch(() => buildDocumentReviewState(0, false));
      return {
        status: "error" as const,
        trackedChangesAffected: 0,
        commentDeleted: false,
        pendingAfter,
        documentState: deriveDocumentState(pendingAfter),
        error: message,
      };
    }
  }

  // -------------------------------------------------------------------------
  // CC selection
  // -------------------------------------------------------------------------

  /** Finds the Content Control for a suggestion using the canonical Stylistic tag. */
  private findCCByTag(
    context: Word.RequestContext,
  ): Word.ContentControlCollection {
    const result = context.document.contentControls.getByTag(
      `stylistic:${this.suggestion.type}:${this.suggestion.id}`,
    );
    result.load("items/tag,items/title");
    return result;
  }

  /** Chooses the best CC candidate when multiple artifacts share the same tag. */
  private selectResolutionContentControl(
    ccs: Word.ContentControl[],
  ): Word.ContentControl | null {
    if (ccs.length === 0) {
      return null;
    }

    const v2Candidate = ccs.find((cc) => {
      const identity = parseReplaceIdentityTitle(cc.title);
      return isValidCompoundReplaceIdentity(identity, this.suggestion);
    });

    return v2Candidate ?? ccs[0] ?? null;
  }

  /** Orders CC candidates so valid compound-v2 artifacts are tried first. */
  private rankResolutionContentControls(
    ccs: Word.ContentControl[],
  ): Word.ContentControl[] {
    return [...ccs].sort((left, right) => {
      const leftValid = isValidCompoundReplaceIdentity(
        parseReplaceIdentityTitle(left.title),
        this.suggestion,
      );
      const rightValid = isValidCompoundReplaceIdentity(
        parseReplaceIdentityTitle(right.title),
        this.suggestion,
      );

      return Number(rightValid) - Number(leftValid);
    });
  }

  // -------------------------------------------------------------------------
  // Observation
  // -------------------------------------------------------------------------

  /**
   * Resolves all tracked changes semantically tied to a suggestion CC.
   *
   * `cc.getTrackedChanges()` can miss one side of a replace operation in real
   * Word documents. To avoid false-success accept/reject flows, this method
   * unions the CC-scoped tracked changes with body-level tracked changes whose
   * ranges overlap the CC range.
   */
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

    console.log(
      `🔎 [ResolveSuggestionCommand] TC sources for CC "${cc.tag}": cc=${ccTrackedChanges.items.length}, ccRange=${rangeTrackedChanges.items.length}, bodyCandidates=${candidateBodyTrackedChanges.length}`,
    );

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

      console.log(
        `🔎 [ResolveSuggestionCommand] body TC candidate #${index + 1} relation=${String(comparisons[index].value)} included=${RESOLUTION_RELATED_RELATIONS.has(comparisons[index].value as string)}`,
      );
    }

    console.log(
      `🔎 [ResolveSuggestionCommand] resolved ${trackedChangesById.size + trackedChangesWithoutId.length} unique tracked changes for CC "${cc.tag}"`,
    );

    return [
      ...Array.from(trackedChangesById.values()),
      ...trackedChangesWithoutId,
    ];
  }

  /** Searches a body or range using the standard fallback strategy. */
  private async searchWithFallback(
    context: Word.RequestContext,
    container: SearchContainer,
    searchText: string,
  ): Promise<Word.Range | null> {
    const searchOptions = { matchCase: true, matchWholeWord: false };

    let results!: Word.RangeCollection;
    if (searchText.length <= 256) {
      results = container.search(searchText, searchOptions);
      results.load("items");
      await context.sync();
    }

    if (searchText.length > 256 || results.items.length === 0) {
      results = container.search(searchText, {
        matchCase: true,
        matchWholeWord: false,
        ignorePunct: true,
        ignoreSpace: true,
      });
      results.load("items");
      await context.sync();
    }

    if (results.items.length > 0) {
      return results.items[0];
    }

    container.load("text");
    await context.sync();

    const fallbackSearchText = findWhitespaceInsensitiveSlice(
      searchText,
      container.text,
    );
    if (!fallbackSearchText) {
      return null;
    }

    results = container.search(fallbackSearchText, searchOptions);
    results.load("items");
    await context.sync();

    return results.items[0] ?? null;
  }

  /** Relocates the operational anchor range persisted in compound-v2 metadata. */
  private async resolveOperationalAnchorRange(
    context: Word.RequestContext,
    identity: ReplaceSuggestionIdentity,
  ): Promise<Word.Range | null> {
    const anchorText = identity.anchorRef?.value?.trim() ?? "";
    if (anchorText.length === 0) {
      console.log(
        "🔎 [ResolveSuggestionCommand] operational anchor missing or empty",
      );
      return null;
    }

    console.log(
      `🔎 [ResolveSuggestionCommand] searching operational anchor (${anchorText.length} chars)`,
    );

    return this.searchWithFallback(
      context,
      context.document.body as unknown as SearchContainer,
      anchorText,
    );
  }

  /**
   * Observes replace suggestion evidence through compound-v2 metadata only.
   */
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

    console.log(
      `🔎 [ResolveSuggestionCommand] observeReplaceSuggestion ccTag="${cc.tag}" hasV2Title=${Boolean((cc as { title?: string }).title?.startsWith(STYLISTIC_IDENTITY_TITLE_PREFIX))}`,
    );

    if (
      (cc as { title?: string }).title?.startsWith(
        STYLISTIC_IDENTITY_TITLE_PREFIX,
      )
    ) {
      if (!isValidCompoundReplaceIdentity(parsedIdentity, this.suggestion)) {
        console.warn(
          `⚠️ [ResolveSuggestionCommand] invalid compound-v2 identity for suggestion "${this.suggestion.id}"`,
          parsedIdentity,
        );
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
        console.log(
          `🔎 [ResolveSuggestionCommand] replace observation succeeded via CC/body sources with ${trackedChanges.length} tracked changes`,
        );
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

        console.log(
          `🔎 [ResolveSuggestionCommand] operational anchor range exposed ${anchorTrackedChanges.items.length} tracked changes`,
        );

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

        console.log(
          `🔎 [ResolveSuggestionCommand] colocated comment range exposed ${commentTrackedChanges.items.length} tracked changes`,
        );

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

    console.warn(
      `⚠️ [ResolveSuggestionCommand] replace suggestion "${this.suggestion.id}" has no compound-v2 title on CC "${cc.tag}"`,
    );
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
    if (isReplaceSuggestion(this.suggestion)) {
      const observation = await this.observeReplaceSuggestion(
        context,
        candidate,
        colocatedComment,
      );
      console.log(
        `🎯 [ResolveSuggestionCommand] replace observation result suggestionId="${this.suggestion.id}" candidateTag="${candidate.tag}": status=${observation.observationStatus}, trackedChanges=${observation.trackedChanges.length}`,
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
    console.log(
      `🎯 [ResolveSuggestionCommand] non-replace observation result suggestionId="${this.suggestion.id}" candidateTag="${candidate.tag}": status=${observationStatus}, trackedChanges=${trackedChanges.length}`,
    );
    return { trackedChanges, observationStatus };
  }

  /** Chooses the best observed candidate by scanning ranked CCs until evidence is conclusive. */
  private async observeResolutionCandidates(
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
      const colocatedComment = await this.findColocatedStylisticComment(
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

  // -------------------------------------------------------------------------
  // Comment management
  // -------------------------------------------------------------------------

  /** Finds the Stylistic comment colocated with the suggestion CC range. */
  private async findColocatedStylisticComment(
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ): Promise<ColocatedCommentContext | null> {
    const comments = context.document.body.getComments();
    comments.load({ select: "authorName,content" });
    await context.sync();

    const stylisticComments = comments.items.filter(isStylisticComment);
    const ccRange = cc.getRange();

    console.log(
      `🔎 [ResolveSuggestionCommand] searching colocated comment for CC "${cc.tag}" among ${stylisticComments.length} Stylistic comments`,
    );

    for (const comment of stylisticComments) {
      const commentRange = comment.getRange();
      const locationResult = commentRange.compareLocationWith(ccRange);
      await context.sync();
      console.log(
        `🔎 [ResolveSuggestionCommand] comment relation to CC "${cc.tag}": ${String(locationResult.value)}`,
      );
      if (OVERLAPPING_RELATIONS.includes(locationResult.value as string)) {
        console.log(
          `🔎 [ResolveSuggestionCommand] found colocated comment for CC "${cc.tag}"`,
        );
        return { comment, range: commentRange };
      }
    }

    console.log(
      `🔎 [ResolveSuggestionCommand] no colocated comment found for CC "${cc.tag}"`,
    );

    return null;
  }

  /** Deletes a previously located Stylistic comment. */
  private async deleteLocatedStylisticComment(
    context: Word.RequestContext,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<boolean> {
    if (!colocatedComment) {
      console.log(
        "🔎 [ResolveSuggestionCommand] no colocated comment to delete",
      );
      return false;
    }

    colocatedComment.comment.delete();
    await context.sync();
    console.log(
      "🔎 [ResolveSuggestionCommand] deleted colocated comment after confirmed resolution",
    );
    return true;
  }

  /**
   * Deletes a colocated Stylistic comment after tracked changes have been resolved.
   *
   * Reject can legitimately invalidate the Word context (proxies, comment ranges)
   * as a side effect of rejecting tracked changes. In that case the `context.sync()`
   * inside the delete throws `GeneralException`. This method tolerates that failure
   * for reject, because the tracked changes were already successfully resolved in Word
   * and the comment will be cleaned up by the next CommentCleanup cycle.
   */
  private async deleteLocatedStylisticCommentAfterResolution(
    context: Word.RequestContext,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<boolean> {
    try {
      return await this.deleteLocatedStylisticComment(
        context,
        colocatedComment,
      );
    } catch (deleteError) {
      if (this.action === "accept") {
        throw deleteError;
      }

      console.warn(
        `⚠️ [ResolveSuggestionCommand] "${this.suggestion.id}": reject comment cleanup failed after successful resolution (will be cleaned up by CommentCleanup): ${deleteError instanceof Error ? deleteError.message : String(deleteError)}`,
      );
      return false;
    }
  }

  // -------------------------------------------------------------------------
  // Resolution actions
  // -------------------------------------------------------------------------

  /** Applies the requested terminal action to all observed tracked changes. */
  private applyTrackedChangeResolution(
    trackedChanges: Word.TrackedChange[],
  ): void {
    for (const tc of trackedChanges) {
      if (this.action === "accept") {
        tc.accept();
      } else {
        tc.reject();
      }
    }

    console.log(
      `🎯 [ResolveSuggestionCommand] executed ${this.action} on ${trackedChanges.length} tracked changes for suggestionId="${this.suggestion.id}"`,
    );
  }

  /** Resolves the comment-only branch by deleting the CC and returning terminal status. */
  private async resolveCommentOnlySuggestion(
    context: Word.RequestContext,
    cc: Word.ContentControl,
    commentDeleted: boolean,
    pendingBefore: DocumentReviewState,
  ): Promise<SuggestionActionResult> {
    cc.delete(true);
    await context.sync();
    const pendingAfter = await inspectDocumentReviewState(context);
    console.log(
      `🗨️ [ResolveSuggestionCommand] "${this.suggestion.id}": comment-only ${this.action}ed, comentario eliminado: ${commentDeleted}`,
    );

    return this.buildResolutionResult(
      this.toResolutionStatus(),
      0,
      commentDeleted,
      pendingBefore,
      pendingAfter,
    );
  }

  /**
   * Deletes the CC anchor after tracked changes were already resolved.
   * Reject can legitimately lose the inserted-side CC as a side effect.
   */
  private async cleanupResolvedSuggestionAnchor(
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ): Promise<void> {
    try {
      cc.delete(true);
      await context.sync();
    } catch (cleanupError) {
      if (this.action === "accept") {
        throw cleanupError;
      }

      console.warn(
        `⚠️ [ResolveSuggestionCommand] "${this.suggestion.id}": reject cleanup skipped after successful resolution: ${cleanupError instanceof Error ? cleanupError.message : String(cleanupError)}`,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Post-resolution state inspection
  // -------------------------------------------------------------------------

  /** Reads pending review state after resolution, tolerating reject-side invalidation. */
  private async inspectDocumentReviewStateAfterResolution(
    context: Word.RequestContext,
    pendingBefore: DocumentReviewState,
  ): Promise<DocumentReviewState> {
    try {
      return await inspectDocumentReviewState(context);
    } catch (error) {
      if (this.action === "accept") {
        throw error;
      }

      console.warn(
        `⚠️ [ResolveSuggestionCommand] "${this.suggestion.id}": reject post-resolution state inspection failed, falling back to pendingBefore snapshot: ${error instanceof Error ? error.message : String(error)}`,
      );
      return pendingBefore;
    }
  }

  // -------------------------------------------------------------------------
  // Result builders
  // -------------------------------------------------------------------------

  /** Maps a resolution action to its terminal success status. */
  private toResolutionStatus(): ResolutionStatus {
    return this.action === "accept"
      ? ("accepted" as const)
      : ("rejected" as const);
  }

  /** Builds a document-aware resolution result. */
  private buildResolutionResult(
    status: SuggestionActionResult["status"],
    trackedChangesAffected: number,
    commentDeleted: boolean,
    pendingBefore: DocumentReviewState,
    pendingAfter: DocumentReviewState,
    error?: string,
  ): SuggestionActionResult {
    const transition = DocumentReviewStateMachine.evaluateTransition(
      pendingBefore,
      pendingAfter,
    );

    return {
      status,
      trackedChangesAffected,
      commentDeleted,
      pendingAfter,
      documentState: transition.to,
      ...(error ? { error } : {}),
    };
  }

  /** Builds a terminal resolution result for identity-loss or unobservable evidence. */
  private async buildObservationFailureResult(
    context: Word.RequestContext,
    status: "identity-lost" | "unobservable",
    pendingBefore: DocumentReviewState,
  ): Promise<SuggestionActionResult> {
    const pendingAfter = await inspectDocumentReviewState(context);
    const error =
      status === "identity-lost"
        ? "La metadata compound-v2 de la sugerencia está incompleta o corrupta."
        : "Word no expuso suficientes tracked changes para confirmar la resolución.";

    return this.buildResolutionResult(
      status,
      0,
      false,
      pendingBefore,
      pendingAfter,
      error,
    );
  }
}
