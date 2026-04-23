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
      console.log(
        `🧹 [SuggestionResolutionCleanup] action=${this.action} suggestionId="${this.suggestionId}" comment=missing`,
      );
      return false;
    }

    console.log(
      `🧹 [SuggestionResolutionCleanup] action=${this.action} suggestionId="${this.suggestionId}" comment=delete-start`,
    );
    colocatedComment.comment.delete();
    await context.sync();
    console.log(
      `🧹 [SuggestionResolutionCleanup] action=${this.action} suggestionId="${this.suggestionId}" comment=delete-done`,
    );
    return true;
  }

  /** Deletes a colocated comment after tracked changes have already been resolved. */
  async deleteLocatedStylisticCommentAfterResolution(
    context: Word.RequestContext,
    colocatedComment: ColocatedCommentContext | null,
  ): Promise<boolean> {
    return this.deleteLocatedStylisticComment(context, colocatedComment);
  }

  /** Deletes the resolved CC anchor as part of atomic cleanup. */
  async cleanupResolvedSuggestionAnchor(
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ): Promise<void> {
    console.log(
      `🧹 [SuggestionResolutionCleanup] action=${this.action} suggestionId="${this.suggestionId}" anchor=delete-start tag="${cc.tag}"`,
    );
    cc.delete(true);
    await context.sync();
    console.log(
      `🧹 [SuggestionResolutionCleanup] action=${this.action} suggestionId="${this.suggestionId}" anchor=delete-done tag="${cc.tag}"`,
    );
  }
}
