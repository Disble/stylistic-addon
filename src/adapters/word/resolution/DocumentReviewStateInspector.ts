import {
  DocumentReviewStateMachine,
  type DocumentReviewUiState,
} from "../../../domain/review/DocumentReviewStateMachine";
import type { DocumentReviewState } from "../../../domain/types";
import { STYLISTIC_TAG_PREFIX } from "../../../infrastructure/config";

/** Creates document-derived review snapshots and reuses them safely. */
export class DocumentReviewStateInspector {
  /** Creates a normalized document-review snapshot. */
  buildDocumentReviewState(
    pendingStylisticArtifacts: number,
    trackChangesActive: boolean,
  ): DocumentReviewState {
    return {
      pendingStylisticArtifacts,
      hasPendingStylisticArtifacts: pendingStylisticArtifacts > 0,
      trackChangesActive,
    };
  }

  /** Builds the safest empty fallback state for outer error recovery. */
  buildEmptyState(): DocumentReviewState {
    return this.buildDocumentReviewState(0, false);
  }

  /** Derives the explicit UI state from a document snapshot. */
  deriveDocumentState(reviewState: DocumentReviewState): DocumentReviewUiState {
    return DocumentReviewStateMachine.deriveState(reviewState);
  }

  /** Reads the authoritative document-derived review state in the current batch. */
  async inspect(context: Word.RequestContext): Promise<DocumentReviewState> {
    const allCCs = context.document.contentControls;
    allCCs.load("items/tag");
    context.document.load("changeTrackingMode");
    await context.sync();

    const pendingStylisticArtifacts = allCCs.items.filter((cc) =>
      cc.tag.startsWith(STYLISTIC_TAG_PREFIX),
    ).length;
    const trackChangesActive =
      context.document.changeTrackingMode !== Word.ChangeTrackingMode.off;

    return this.buildDocumentReviewState(
      pendingStylisticArtifacts,
      trackChangesActive,
    );
  }

  /** Reads pending review state after resolution, tolerating reject-side invalidation. */
  async inspectAfterResolution(
    context: Word.RequestContext,
    pendingBefore: DocumentReviewState,
    action: "accept" | "reject",
    suggestionId: string,
  ): Promise<DocumentReviewState> {
    try {
      return await this.inspect(context);
    } catch (error) {
      if (action === "accept") {
        throw error;
      }

      console.warn(
        `⚠️ [DocumentReviewStateInspector] "${suggestionId}": reject post-resolution state inspection failed, falling back to pendingBefore snapshot: ${error instanceof Error ? error.message : String(error)}`,
      );
      return pendingBefore;
    }
  }
}
