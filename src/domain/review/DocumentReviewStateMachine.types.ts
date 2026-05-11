/**
 * Types owned by `DocumentReviewStateMachine` and review-state mediation.
 */

/**
 * Document-derived review state for Stylistic artifacts currently materialized
 * in Word.
 */
export interface DocumentReviewState {
  /** Number of pending Stylistic artifacts still active in the document. */
  pendingStylisticArtifacts: number;

  /** Convenience boolean derived from `pendingStylisticArtifacts > 0`. */
  hasPendingStylisticArtifacts: boolean;

  /** Whether Word Track Changes is currently active for the document. */
  trackChangesActive: boolean;
}

/** Explicit document-review states consumed by the taskpane/workflows. */
export type DocumentReviewUiState = "idle" | "pending-review" | "ready-to-disable-track-changes";

/** Transition metadata exposed to workflow/application callers. */
export interface DocumentReviewTransition {
  /** Previous UI state. */
  from: DocumentReviewUiState;

  /** Next UI state. */
  to: DocumentReviewUiState;
}

/** Taskpane-facing review state derived by the explicit review mediator. */
export interface ReviewTaskpaneState {
  /** Explicit document-review UI state currently exposed to the user. */
  documentState: DocumentReviewUiState;

  /** Whether the taskpane should expose the final Track Changes deactivation CTA. */
  showDisableTrackChangesCta: boolean;

  /** Whether the cleanup section should currently be visible in the taskpane. */
  showCleanupSection: boolean;
}
