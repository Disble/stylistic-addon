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
 * Algorithm (4 Word API syncs):
 * 1. Load all Stylistic comments and tracked changes with properties.
 * 2. Get document ranges for each comment and tracked change.
 * 3. Compare every comment range against every TC range (spatial matrix).
 * 4. Delete comments with no overlapping TC; keep the rest.
 *
 * Never touches comments authored by anyone other than "Stylistic".
 *
 * @module CommentCleanup
 */

/** All `LocationRelation` values that indicate spatial overlap. */
const OVERLAPPING_RELATIONS: string[] = [
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
 *
 * @returns Counts of deleted and kept Stylistic comments.
 */
export async function cleanupResolvedComments(): Promise<{
  deleted: number;
  kept: number;
}> {
  console.log("🧽 [CommentCleanup] Iniciando limpieza de comentarios resueltos...");

  return Word.run(async (context) => {
    // Sync 1: load collections with filtering properties
    const tracked = context.document.body.getTrackedChanges();
    const comments = context.document.body.getComments();
    tracked.load({ select: "author,type" });
    comments.load({ select: "authorName" });
    await context.sync();

    const stylisticComments = comments.items.filter((c) => c.authorName === "Stylistic");
    const stylisticTCs = tracked.items.filter((tc) => tc.author === "Stylistic");
    console.log(
      `🧽 [CommentCleanup] ${stylisticComments.length} comentarios y ${stylisticTCs.length} TCs de Stylistic`
    );

    if (stylisticComments.length === 0) {
      return { deleted: 0, kept: 0 };
    }

    // Short-circuit: all TCs resolved → delete all Stylistic comments
    if (stylisticTCs.length === 0) {
      for (const comment of stylisticComments) {
        comment.delete();
      }
      await context.sync();
      return { deleted: stylisticComments.length, kept: 0 };
    }

    // Sync 2: get document ranges
    const commentRanges = stylisticComments.map((c) => c.getRange());
    const tcRanges = stylisticTCs.map((tc) => tc.getRange());
    await context.sync();

    // Sync 3: build spatial comparison matrix
    const comparisons: OfficeExtension.ClientResult<Word.LocationRelation>[][] = [];
    for (let i = 0; i < commentRanges.length; i++) {
      comparisons[i] = [];
      for (let j = 0; j < tcRanges.length; j++) {
        comparisons[i][j] = commentRanges[i].compareLocationWith(tcRanges[j]);
      }
    }
    await context.sync();

    // Evaluate colocation and schedule deletes
    let deleted = 0;
    let kept = 0;

    for (let i = 0; i < stylisticComments.length; i++) {
      const hasColocatedTC = tcRanges.some((_, j) =>
        OVERLAPPING_RELATIONS.includes(comparisons[i][j].value as string)
      );

      if (hasColocatedTC) {
        kept++;
      } else {
        stylisticComments[i].delete();
        deleted++;
      }
    }

    // Sync 4: execute deletes
    await context.sync();
    console.log(`🧽 [CommentCleanup] ${deleted} eliminados, ${kept} conservados`);
    return { deleted, kept };
  });
}
