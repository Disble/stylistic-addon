import type { SuggestionResolutionWarning } from "../../../domain/types";
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
  ): Promise<{ deleted: boolean; warning?: SuggestionResolutionWarning }> {
    try {
      const deleted = await this.deleteLocatedStylisticComment(
        context,
        colocatedComment,
      );
      return { deleted };
    } catch (deleteError) {
      const message =
        deleteError instanceof Error
          ? deleteError.message
          : String(deleteError);

      console.warn(
        `⚠️ [SuggestionResolutionCleanup] "${this.suggestionId}": ${this.action} comment cleanup failed after semantic resolution (will be cleaned up later): ${message}`,
      );
      return {
        deleted: false,
        warning: {
          code: "cleanup-failed",
          phase: "cleanup",
          message,
        },
      };
    }
  }

  /** Deletes the resolved CC anchor, tolerating reject-side invalidation. */
  async cleanupResolvedSuggestionAnchor(
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ): Promise<SuggestionResolutionWarning | undefined> {
    try {
      cc.delete(true);
      await context.sync();
    } catch (cleanupError) {
      const message =
        cleanupError instanceof Error
          ? cleanupError.message
          : String(cleanupError);

      console.warn(
        `⚠️ [SuggestionResolutionCleanup] "${this.suggestionId}": ${this.action} anchor cleanup skipped after semantic resolution: ${message}`,
      );
      return {
        code: "cleanup-failed",
        phase: "cleanup",
        message,
      };
    }
  }
}
