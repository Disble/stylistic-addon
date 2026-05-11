import type { DocumentReviewUiState } from "./DocumentReviewStateMachine.types";

/** Valid transitions for the explicit document-review UI machine. */
export const DOCUMENT_REVIEW_TRANSITIONS: Record<DocumentReviewUiState, DocumentReviewUiState[]> = {
  idle: ["idle", "pending-review", "ready-to-disable-track-changes"],
  "pending-review": ["pending-review", "ready-to-disable-track-changes", "idle"],
  "ready-to-disable-track-changes": ["ready-to-disable-track-changes", "pending-review", "idle"],
};
