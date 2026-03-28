/* global Word, console, OfficeExtension */

/**
 * Comment cleanup — Range Colocation pattern for orphaned comment removal.
 *
 * Deletes Stylistic comments whose tracked changes have been resolved
 * (accepted or rejected by the user). Uses `Range.compareLocationWith()` to
 * determine whether a comment is still spatially anchored to a pending
 * tracked change. The document itself is the source of truth — no in-memory
 * registry is needed and no cross-session state is required.
 *
 * Algorithm:
 * 0. Pre-filter: load all Content Controls and identify those owned by active
 *    comment-only suggestions (tag prefix `stylistic:comment-only:`). Comments
 *    anchored to those CCs are excluded from the TC-colocation check — they
 *    are managed by their own lifecycle and must not be deleted here.
 * 1. Load all Stylistic comments and tracked changes with properties.
 * 2. Short-circuit only for "track-change" comments (those NOT owned by a
 *    comment-only CC). If all TCs are resolved, delete only those comments.
 * 3. Get document ranges for each remaining comment and tracked change.
 * 4. Compare every comment range against every TC range (spatial matrix).
 * 5. Delete track-change comments with no overlapping TC; keep the rest.
 *
 * Backward compatibility: CC tags without the `stylistic:` prefix (legacy bare
 * IDs) are treated as `track-change` — they are not comment-only.
 *
 * Never touches comments authored by anyone other than "Stylistic".
 *
 * @module CommentCleanup
 */

/** Tag prefix that identifies comment-only suggestion Content Controls. */
export const COMMENT_ONLY_TAG_PREFIX = "stylistic:comment-only:";

/** All `LocationRelation` values that indicate spatial overlap. */
export const OVERLAPPING_RELATIONS: string[] = [
  "Equal",
  "Contains",
  "ContainsStart",
  "ContainsEnd",
  "Inside",
  "InsideStart",
  "InsideEnd",
  "OverlapsBefore",
  "OverlapsAfter",
];

/**
 * Deletes Stylistic comments whose tracked changes have been resolved.
 * Comments owned by active comment-only Content Controls are never deleted
 * by this function — they are managed by `resolveSuggestion()`.
 *
 * @returns Counts of deleted and kept Stylistic comments.
 */
export async function cleanupResolvedComments(): Promise<{
  deleted: number;
  kept: number;
}> {
  console.log("🧽 [CommentCleanup] Iniciando limpieza de comentarios resueltos...");

  return Word.run(async (context) => {
    // Sync 1: load collections with filtering properties + all Content Controls
    const tracked = context.document.body.getTrackedChanges();
    const comments = context.document.body.getComments();
    const allCCs = context.document.contentControls;
    tracked.load({ select: "author,type" });
    comments.load({ select: "authorName" });
    allCCs.load({ select: "tag" });
    await context.sync();

    const stylisticComments = comments.items.filter((c) => c.authorName === "Stylistic");
    const stylisticTCs = tracked.items.filter((tc) => tc.author === "Stylistic");

    // Pre-filter: find all active comment-only Content Controls (JS-side prefix filter,
    // because Office.js getByTag() is exact-match only — no prefix query available).
    const commentOnlyCCs = allCCs.items.filter((cc) =>
      cc.tag.startsWith(COMMENT_ONLY_TAG_PREFIX)
    );

    console.log(
      `🧽 [CommentCleanup] ${stylisticComments.length} comentarios, ${stylisticTCs.length} TCs, ` +
        `${commentOnlyCCs.length} CCs comment-only activos`
    );

    if (stylisticComments.length === 0) {
      return { deleted: 0, kept: 0 };
    }

    // Sync 2: get ranges for all comment-only CCs so we can match comments to them
    const commentOnlyCCRanges = commentOnlyCCs.map((cc) => cc.getRange());
    const commentRangesForCCCheck = stylisticComments.map((c) => c.getRange());
    await context.sync();

    // Sync 3: build colocation matrix between comments and comment-only CC ranges
    let commentOnlyCommentIndices = new Set<number>();

    if (commentOnlyCCs.length > 0) {
      const ccComparisons: OfficeExtension.ClientResult<Word.LocationRelation>[][] = [];
      for (let i = 0; i < stylisticComments.length; i++) {
        ccComparisons[i] = [];
        for (let j = 0; j < commentOnlyCCRanges.length; j++) {
          ccComparisons[i][j] = commentRangesForCCCheck[i].compareLocationWith(
            commentOnlyCCRanges[j]
          );
        }
      }
      await context.sync();

      // Identify which Stylistic comments are anchored to a comment-only CC
      for (let i = 0; i < stylisticComments.length; i++) {
        const isOwnedByCommentOnlyCC = commentOnlyCCRanges.some((_, j) =>
          OVERLAPPING_RELATIONS.includes(ccComparisons[i][j].value as string)
        );
        if (isOwnedByCommentOnlyCC) {
          commentOnlyCommentIndices.add(i);
        }
      }
    }

    console.log(
      `🧽 [CommentCleanup] ${commentOnlyCommentIndices.size} comentario(s) protegidos por CCs comment-only`
    );

    // Separate track-change comments (subject to TC-colocation check) from
    // comment-only comments (managed by resolveSuggestion — skip here).
    const trackChangeCommentIndices = stylisticComments
      .map((_, i) => i)
      .filter((i) => !commentOnlyCommentIndices.has(i));

    const trackChangeComments = trackChangeCommentIndices.map((i) => stylisticComments[i]);

    if (trackChangeComments.length === 0) {
      console.log("🧽 [CommentCleanup] Sin comentarios track-change que limpiar");
      return { deleted: 0, kept: commentOnlyCommentIndices.size };
    }

    // Short-circuit: all TCs resolved → delete only track-change comments
    if (stylisticTCs.length === 0) {
      for (const comment of trackChangeComments) {
        comment.delete();
      }
      await context.sync();
      console.log(
        `🧽 [CommentCleanup] ${trackChangeComments.length} eliminados (short-circuit), ` +
          `${commentOnlyCommentIndices.size} conservados (comment-only)`
      );
      return {
        deleted: trackChangeComments.length,
        kept: commentOnlyCommentIndices.size,
      };
    }

    // Sync 4: get document ranges for track-change comments and TCs
    const tcCommentRanges = trackChangeComments.map((c) => c.getRange());
    const tcRanges = stylisticTCs.map((tc) => tc.getRange());
    await context.sync();

    // Sync 5: build spatial comparison matrix (track-change comments × TCs)
    const comparisons: OfficeExtension.ClientResult<Word.LocationRelation>[][] = [];
    for (let i = 0; i < tcCommentRanges.length; i++) {
      comparisons[i] = [];
      for (let j = 0; j < tcRanges.length; j++) {
        comparisons[i][j] = tcCommentRanges[i].compareLocationWith(tcRanges[j]);
      }
    }
    await context.sync();

    // Evaluate colocation and schedule deletes for track-change comments only
    let deleted = 0;
    let kept = commentOnlyCommentIndices.size; // comment-only comments always kept here

    for (let i = 0; i < trackChangeComments.length; i++) {
      const hasColocatedTC = tcRanges.some((_, j) =>
        OVERLAPPING_RELATIONS.includes(comparisons[i][j].value as string)
      );

      if (hasColocatedTC) {
        kept++;
      } else {
        trackChangeComments[i].delete();
        deleted++;
      }
    }

    // Sync 6: execute deletes
    await context.sync();
    console.log(`🧽 [CommentCleanup] ${deleted} eliminados, ${kept} conservados`);
    return { deleted, kept };
  });
}
