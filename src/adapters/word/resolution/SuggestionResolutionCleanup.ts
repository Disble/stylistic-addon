import type { ColocatedCommentContext } from "./ResolutionContext";

/** Owns cleanup policy after tracked changes were already resolved. */
export class SuggestionResolutionCleanup {
  constructor(
    private readonly suggestionId: string,
    private readonly action: "accept" | "reject",
  ) {}

  /** Deletes a previously located Stylistic comment. */
  async deleteLocatedStylisticComment(
    context: Word.RequestContext,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<boolean> {
    if (!colocatedComment) {
      return false;
    }

    colocatedComment.comment.delete();
    await context.sync();
    return true;
  }

  /** Deletes a colocated comment after tracked changes have already been resolved. */
  async deleteLocatedStylisticCommentAfterResolution(
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
        `⚠️ [SuggestionResolutionCleanup] "${this.suggestionId}": reject comment cleanup failed after successful resolution (will be cleaned up by CommentCleanup): ${deleteError instanceof Error ? deleteError.message : String(deleteError)}`,
      );
      return false;
    }
  }

  /** Deletes the resolved CC anchor, tolerating reject-side invalidation. */
  async cleanupResolvedSuggestionAnchor(
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
        `⚠️ [SuggestionResolutionCleanup] "${this.suggestionId}": reject cleanup skipped after successful resolution: ${cleanupError instanceof Error ? cleanupError.message : String(cleanupError)}`,
      );
    }
  }
}
