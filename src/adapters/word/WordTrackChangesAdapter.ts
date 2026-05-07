/* global Word */

import { DocumentReviewStateMachine } from "../../domain/review/DocumentReviewStateMachine";
import type {
  DocumentReviewState,
  DocumentReviewUiState,
} from "../../domain/review/DocumentReviewStateMachine.types";
import { STYLISTIC_TAG_PREFIX } from "../../infrastructure/config";

/** Owns document review state and Track Changes lifecycle concerns. */
export class WordTrackChangesAdapter {
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

  /** Derives the explicit document-review UI state from a document snapshot. */
  deriveDocumentState(reviewState: DocumentReviewState): DocumentReviewUiState {
    return DocumentReviewStateMachine.deriveState(reviewState);
  }

  /** Returns the current Track Changes activation state for the document. */
  async loadTrackChangesActive(context: Word.RequestContext): Promise<boolean> {
    context.document.load("changeTrackingMode");
    await context.sync();
    return context.document.changeTrackingMode !== Word.ChangeTrackingMode.off;
  }

  /** Reads the authoritative document-derived review state in the current batch. */
  async inspectDocumentReviewState(context: Word.RequestContext): Promise<DocumentReviewState> {
    const allCCs = context.document.contentControls;
    allCCs.load("items/tag");
    context.document.load("changeTrackingMode");
    await context.sync();

    const pendingStylisticArtifacts = allCCs.items.filter((cc) =>
      cc.tag.startsWith(STYLISTIC_TAG_PREFIX)
    ).length;
    const trackChangesActive = context.document.changeTrackingMode !== Word.ChangeTrackingMode.off;

    return this.buildDocumentReviewState(pendingStylisticArtifacts, trackChangesActive);
  }

  /** Enables Track Changes once, lazily, before the first real insertion. */
  async ensureTrackChangesActive(context: Word.RequestContext): Promise<boolean> {
    const alreadyActive = await this.loadTrackChangesActive(context);
    if (alreadyActive) {
      return false;
    }

    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();
    return true;
  }

  /** Returns the current document-derived Stylistic review state. */
  async getDocumentReviewState(): Promise<DocumentReviewState> {
    return Word.run((context) => this.inspectDocumentReviewState(context));
  }

  /** Disables Track Changes only when the user explicitly requests it. */
  async disableTrackChanges(): Promise<void> {
    await Word.run(async (context) => {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    });
  }
}
