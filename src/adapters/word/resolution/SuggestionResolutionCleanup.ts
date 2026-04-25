import type { ColocatedCommentContext } from "./ResolutionContext";

/** Owns cleanup policy after tracked changes were already resolved. */
export class SuggestionResolutionCleanup {
  constructor(
    private readonly suggestionId: string,
    private readonly action: "accept" | "reject",
  ) {}

  /**
   * Deletes a previously located Stylistic comment.
   *
   * Treats `GeneralException` from `comment.delete()` / `context.sync()` as a
   * soft success: this almost always means the host already invalidated the
   * comment proxy as a side-effect of the preceding tracked-change resolution
   * (typical when rejecting a replace removes the surrounding CC and the
   * colocated comment proxy points to a now-orphaned anchor). The user-visible
   * outcome is identical — the comment is no longer attached to the document —
   * so we MUST NOT surface this as a workflow error and leave the taskpane
   * card stuck in "pending with error". Other host errors are propagated so
   * legitimate cleanup failures still abort the workflow.
   */
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
    try {
      colocatedComment.comment.delete();
      await context.sync();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const code = (error as { code?: string } | null)?.code ?? "";
      const isGeneralException =
        code === "GeneralException" || message.includes("GeneralException");

      if (isGeneralException) {
        console.warn(
          `🧹 [SuggestionResolutionCleanup] action=${this.action} suggestionId="${this.suggestionId}" comment=delete-soft-success (host invalidated comment proxy after prior mutation: ${message})`,
        );
        return true;
      }

      throw error;
    }
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
}
