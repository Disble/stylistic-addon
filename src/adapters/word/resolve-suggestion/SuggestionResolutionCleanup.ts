import {
  STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../../infrastructure/config";
import { stringifyUnknownError } from "../../../infrastructure/errors/UnknownError.helpers";
import type { ColocatedCommentContext } from "./ResolutionContext";
import type { ResolvedTrackChangeMetadataCleanupResult } from "./SuggestionResolutionCleanup.types";

/**
 * Owns cleanup policy after tracked changes were already resolved.
 *
 * Real Word validation showed that resolving native tracked changes is not
 * enough: inserted-side and operational-wrapper Content Controls can remain in
 * OOXML and later collide with new suggestions. Therefore, track-change cleanup
 * must delete suggestion-owned metadata explicitly after semantic resolution.
 */
export class SuggestionResolutionCleanup {
  constructor(
    private readonly suggestionId: string,
    private readonly action: "accept" | "reject"
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
    colocatedComment: ColocatedCommentContext | null
  ): Promise<boolean> {
    if (!colocatedComment) {
      console.log(
        `🧹 [SuggestionResolutionCleanup] action=${this.action} suggestionId="${this.suggestionId}" comment=missing`
      );
      return false;
    }

    console.log(
      `🧹 [SuggestionResolutionCleanup] action=${this.action} suggestionId="${this.suggestionId}" comment=delete-start`
    );
    try {
      colocatedComment.comment.delete();
      await context.sync();
    } catch (error) {
      const message = this.stringifyError(error);
      const code = (error as { code?: string } | null)?.code ?? "";
      const isGeneralException =
        code === "GeneralException" || message.includes("GeneralException");

      if (isGeneralException) {
        console.warn(
          `🧹 [SuggestionResolutionCleanup] action=${this.action} suggestionId="${this.suggestionId}" comment=delete-soft-success (host invalidated comment proxy after prior mutation: ${message})`
        );
        return true;
      }

      throw error;
    }
    console.log(
      `🧹 [SuggestionResolutionCleanup] action=${this.action} suggestionId="${this.suggestionId}" comment=delete-done`
    );
    return true;
  }

  /** Deletes a colocated comment after tracked changes have already been resolved. */
  async deleteLocatedStylisticCommentAfterResolution(
    context: Word.RequestContext,
    colocatedComment: ColocatedCommentContext | null
  ): Promise<boolean> {
    return this.deleteLocatedStylisticComment(context, colocatedComment);
  }

  /**
   * Deletes resolved track-change metadata after native tracked changes already
   * reached their terminal state.
   *
   * The deletion is intentionally re-located by exact tags instead of using stale
   * proxies captured before `acceptAll()` / `rejectAll()`: Word can invalidate or
   * rematerialize those proxies during native revision resolution. Track Changes
   * is temporarily disabled only for this housekeeping mutation so Word does not
   * preserve the metadata deletion as a new pending revision in OOXML; the
   * user's previous tracking mode is restored immediately afterward.
   */
  async deleteResolvedTrackChangeMetadata(
    context: Word.RequestContext
  ): Promise<ResolvedTrackChangeMetadataCleanupResult> {
    return this.runWithTrackChangesDisabled(context, async () => {
      const targetTags = new Set([this.buildOperationalWrapperTag(), this.buildInsertedSideTag()]);
      const deletedContentControls: string[] = [];
      const failedContentControls: Array<{ tag: string; error: string }> = [];

      const contentControls = context.document.contentControls;
      contentControls.load("items/tag");
      await context.sync();

      for (const contentControl of contentControls.items) {
        if (!targetTags.has(contentControl.tag)) {
          continue;
        }

        try {
          contentControl.delete(true);
          deletedContentControls.push(contentControl.tag);
        } catch (error) {
          failedContentControls.push({
            tag: contentControl.tag,
            error: this.stringifyError(error),
          });
        }
      }

      await context.sync();

      return { deletedContentControls, failedContentControls };
    });
  }

  /** Builds the exact inserted-side tag persisted when the suggestion was applied. */
  private buildInsertedSideTag(): string {
    return `${STYLISTIC_TAG_PREFIX}track-change:${this.suggestionId}`;
  }

  /** Builds the exact operational wrapper tag persisted when the suggestion was applied. */
  private buildOperationalWrapperTag(): string {
    return `${STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX}${this.suggestionId}`;
  }

  /**
   * Runs cleanup with Track Changes disabled and restores the user's previous
   * mode even when the cleanup operation fails.
   */
  private async runWithTrackChangesDisabled<T>(
    context: Word.RequestContext,
    operation: () => Promise<T>
  ): Promise<T> {
    context.document.load("changeTrackingMode");
    await context.sync();

    const previousTrackingMode = context.document.changeTrackingMode;
    const mustDisableTracking = previousTrackingMode !== Word.ChangeTrackingMode.off;

    if (mustDisableTracking) {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    }

    try {
      return await operation();
    } finally {
      if (mustDisableTracking) {
        context.document.changeTrackingMode = previousTrackingMode;
        await context.sync();
      }
    }
  }

  /** Converts unknown host exceptions to stable diagnostics. */
  private stringifyError(error: unknown): string {
    return stringifyUnknownError(error);
  }
}
