import { DocumentReviewStateMachine } from "../../../domain/review/DocumentReviewStateMachine";
import type {
  DocumentReviewState,
  DocumentReviewUiState,
} from "../../../domain/review/DocumentReviewStateMachine.types";
import {
  STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../../infrastructure/config";

/** Creates document-derived review snapshots and reuses them safely. */
export class DocumentReviewStateInspector {
  /** Creates a normalized document-review snapshot. */
  buildDocumentReviewState(
    pendingStylisticArtifacts: number,
    trackChangesActive: boolean
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

    const pendingStylisticArtifacts = allCCs.items.filter(
      (cc) =>
        cc.tag.startsWith(STYLISTIC_TAG_PREFIX) &&
        !cc.tag.startsWith(STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX)
    ).length;
    const trackChangesActive = context.document.changeTrackingMode !== Word.ChangeTrackingMode.off;

    return this.buildDocumentReviewState(pendingStylisticArtifacts, trackChangesActive);
  }

  /** Reads pending review state after resolution and propagates host failures. */
  async inspectAfterResolution(context: Word.RequestContext): Promise<DocumentReviewState> {
    return this.inspect(context);
  }
}
