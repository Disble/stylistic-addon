/* global Word, console, OfficeExtension */

import { isStylisticComment } from "../StylisticCommentBuilder";

/**
 * Comment cleanup — Range Colocation pattern for orphaned comment removal.
 *
 * Deletes Stylistic comments that are no longer anchored to an active
 * Stylistic suggestion Content Control. Uses `Range.compareLocationWith()` to
 * determine whether a comment is still spatially colocated with a
 * `stylistic:` tagged Content Control. The document itself is
 * the source of truth — no in-memory registry is needed and no cross-session
 * state is required.
 *
 * Algorithm:
 * 1. Load all Stylistic comments and all Content Controls.
 * 2. Filter CCs by the `stylistic:` tag prefix.
 * 3. Short-circuit: if no such CCs exist, delete all Stylistic comments.
 * 4. Get document ranges for each comment and each active CC.
 * 5. Compare every comment range against every CC range (spatial matrix).
 * 6. Delete comments with no overlapping CC; keep the rest.
 *
 * Never touches comments authored by anyone other than "Stylistic".
 *
 * @module CommentCleanup
 */

/** Tag prefix that identifies any active Stylistic suggestion Content Control. */
export const STYLISTIC_TAG_PREFIX = "stylistic:";

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
 * Deletes Stylistic comments that are no longer colocated with an active
 * `stylistic:` Content Control.
 *
 * @returns Counts of deleted and kept Stylistic comments.
 */
export async function getCleanupPreview(): Promise<{
  deletable: number;
  kept: number;
}> {
  return Word.run(async (context) => {
    // Sync 1: load collections
    const comments = context.document.body.getComments();
    const allCCs = context.document.contentControls;
    comments.load({ select: "authorName,content" });
    allCCs.load({ select: "tag" });
    await context.sync();

    const stylisticComments = comments.items.filter(isStylisticComment);

    // Pre-filter: find all active Stylistic Content Controls (JS-side prefix
    // filter — Office.js getByTag() is exact-match only, no prefix query).
    const stylisticCCs = allCCs.items.filter((cc) =>
      cc.tag.startsWith(STYLISTIC_TAG_PREFIX),
    );

    console.log(
      `🧽 [CommentCleanup] ${stylisticComments.length} comentarios Stylistic, ` +
        `${stylisticCCs.length} CCs Stylistic activos`,
    );

    if (stylisticComments.length === 0) {
      return { deletable: 0, kept: 0 };
    }

    // Sync 2: get ranges for comments and CCs
    const commentRanges = stylisticComments.map((c) => c.getRange());
    const ccRanges = stylisticCCs.map((cc) => cc.getRange());
    await context.sync();

    // Short-circuit: no active Stylistic CCs → delete all Stylistic comments
    if (stylisticCCs.length === 0) {
      return { deletable: stylisticComments.length, kept: 0 };
    }

    // Sync 3: build spatial comparison matrix (comments × CCs)
    const comparisons: OfficeExtension.ClientResult<Word.LocationRelation>[][] =
      [];
    for (let i = 0; i < commentRanges.length; i++) {
      comparisons[i] = [];
      for (let j = 0; j < ccRanges.length; j++) {
        comparisons[i][j] = commentRanges[i].compareLocationWith(ccRanges[j]);
      }
    }
    await context.sync();

    // Evaluate colocation: keep comments overlapping a CC; delete orphans
    let deletable = 0;
    let kept = 0;

    for (let i = 0; i < stylisticComments.length; i++) {
      const hasColocatedCC = ccRanges.some((_, j) =>
        OVERLAPPING_RELATIONS.includes(comparisons[i][j].value as string),
      );

      if (hasColocatedCC) {
        kept++;
      } else {
        deletable++;
      }
    }

    return { deletable, kept };
  });
}

/**
 * Deletes Stylistic comments that are no longer colocated with an active
 * `stylistic:` Content Control.
 *
 * @returns Counts of deleted and kept Stylistic comments.
 */
export async function cleanupResolvedComments(): Promise<{
  deleted: number;
  kept: number;
}> {
  console.log(
    "🧽 [CommentCleanup] Iniciando limpieza de comentarios resueltos...",
  );

  return Word.run(async (context) => {
    // Sync 1: load collections
    const comments = context.document.body.getComments();
    const allCCs = context.document.contentControls;
    comments.load({ select: "authorName,content" });
    allCCs.load({ select: "tag" });
    await context.sync();

    const stylisticComments = comments.items.filter(isStylisticComment);

    const stylisticCCs = allCCs.items.filter((cc) =>
      cc.tag.startsWith(STYLISTIC_TAG_PREFIX),
    );

    console.log(
      `🧽 [CommentCleanup] ${stylisticComments.length} comentarios Stylistic, ` +
        `${stylisticCCs.length} CCs Stylistic activos`,
    );

    if (stylisticComments.length === 0) {
      return { deleted: 0, kept: 0 };
    }

    // Sync 2: get ranges for comments and CCs
    const commentRanges = stylisticComments.map((c) => c.getRange());
    const ccRanges = stylisticCCs.map((cc) => cc.getRange());
    await context.sync();

    // Short-circuit: no active Stylistic CCs → delete all Stylistic comments
    if (stylisticCCs.length === 0) {
      for (const comment of stylisticComments) {
        comment.delete();
      }
      await context.sync();
      console.log(
        `🧽 [CommentCleanup] ${stylisticComments.length} eliminados (sin CCs activos)`,
      );
      return { deleted: stylisticComments.length, kept: 0 };
    }

    // Sync 3: build spatial comparison matrix (comments × CCs)
    const comparisons: OfficeExtension.ClientResult<Word.LocationRelation>[][] =
      [];
    for (let i = 0; i < commentRanges.length; i++) {
      comparisons[i] = [];
      for (let j = 0; j < ccRanges.length; j++) {
        comparisons[i][j] = commentRanges[i].compareLocationWith(ccRanges[j]);
      }
    }
    await context.sync();

    // Evaluate colocation: keep comments overlapping a CC; delete orphans
    let deleted = 0;
    let kept = 0;

    for (let i = 0; i < stylisticComments.length; i++) {
      const hasColocatedCC = ccRanges.some((_, j) =>
        OVERLAPPING_RELATIONS.includes(comparisons[i][j].value as string),
      );

      if (hasColocatedCC) {
        kept++;
      } else {
        stylisticComments[i].delete();
        deleted++;
      }
    }

    // Sync 4: execute deletes
    await context.sync();
    console.log(
      `🧽 [CommentCleanup] ${deleted} eliminados, ${kept} conservados`,
    );
    return { deleted, kept };
  });
}
