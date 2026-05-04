/* global console */

/**
 * Document Review State Machine — centralizes taskpane/UI semantics derived from
 * the Word document review snapshot.
 *
 * Important distinction:
 * - `DocumentReviewState` is still the authoritative snapshot read from Word.
 * - This machine translates that snapshot into explicit UI/workflow states so
 *   the taskpane stops scattering boolean checks across multiple layers.
 *
 * The machine intentionally models only the states that matter for frontend
 * review consistency:
 * - `idle`                         → no pending Stylistic artifacts, Track Changes off
 * - `pending-review`               → at least one pending Stylistic artifact remains
 * - `ready-to-disable-track-changes` → zero pending Stylistic artifacts, Track Changes still on
 *
 * @module DocumentReviewStateMachine
 */

import type {
  DocumentReviewState,
  DocumentReviewTransition,
  DocumentReviewUiState,
} from "./DocumentReviewStateMachine.types";

export type {
  DocumentReviewState,
  DocumentReviewTransition,
  DocumentReviewUiState,
  ReviewTaskpaneState,
} from "./DocumentReviewStateMachine.types";

const TRANSITIONS: Record<DocumentReviewUiState, DocumentReviewUiState[]> = {
  idle: ["idle", "pending-review", "ready-to-disable-track-changes"],
  "pending-review": ["pending-review", "ready-to-disable-track-changes", "idle"],
  "ready-to-disable-track-changes": ["ready-to-disable-track-changes", "pending-review", "idle"],
};

export class InvalidDocumentReviewTransitionError extends Error {
  constructor(
    from: DocumentReviewUiState,
    to: DocumentReviewUiState,
    allowed: DocumentReviewUiState[]
  ) {
    super(
      `[DocumentReviewStateMachine] Invalid transition: "${from}" → "${to}". ` +
        `Allowed: [${allowed.join(", ")}]`
    );
    this.name = "InvalidDocumentReviewTransitionError";
  }
}

/**
 * State machine that owns document-review UI semantics.
 */
export class DocumentReviewStateMachine {
  private current: DocumentReviewUiState = "idle";

  constructor(initial?: DocumentReviewState) {
    if (initial) {
      this.current = DocumentReviewStateMachine.deriveState(initial);
    }
  }

  /** Current document-review UI state. */
  get state(): DocumentReviewUiState {
    return this.current;
  }

  /** `true` when the explicit post-review Track Changes CTA should be visible. */
  get shouldShowDisableTrackChangesCta(): boolean {
    return this.current === "ready-to-disable-track-changes";
  }

  /**
   * Derives the explicit UI state from the authoritative document snapshot.
   */
  static deriveState(reviewState: DocumentReviewState): DocumentReviewUiState {
    if (reviewState.hasPendingStylisticArtifacts) {
      return "pending-review";
    }

    if (reviewState.trackChangesActive) {
      return "ready-to-disable-track-changes";
    }

    return "idle";
  }

  /**
   * Evaluates the semantic transition between two document snapshots.
   */
  static evaluateTransition(
    before: DocumentReviewState,
    after: DocumentReviewState
  ): DocumentReviewTransition {
    const machine = new DocumentReviewStateMachine(before);
    const from = machine.state;
    machine.syncFromDocument(after);

    return {
      from,
      to: machine.state,
    };
  }

  /**
   * Rehydrates the machine from the current authoritative document snapshot.
   */
  syncFromDocument(reviewState: DocumentReviewState): void {
    this.transitionTo(DocumentReviewStateMachine.deriveState(reviewState));
  }

  /**
   * Marks the taskpane as having active pending review work after successful
   * application of at least one suggestion.
   */
  markPendingReview(): void {
    this.transitionTo("pending-review");
  }

  /**
   * Applies the explicit user action of disabling Track Changes after zero
   * pending Stylistic artifacts remain.
   */
  disableTrackChanges(): void {
    if (this.current === "idle") {
      return;
    }

    if (this.current !== "ready-to-disable-track-changes") {
      throw new InvalidDocumentReviewTransitionError(this.current, "idle", [
        "ready-to-disable-track-changes",
      ]);
    }

    this.transitionTo("idle");
  }

  /**
   * Resets the machine to `idle`.
   */
  reset(): void {
    console.log(`🔄 [DocumentReviewStateMachine] reset → idle (was: ${this.current})`);
    this.current = "idle";
  }

  /**
   * Applies one validated state transition.
   */
  private transitionTo(next: DocumentReviewUiState): void {
    if (!TRANSITIONS[this.current].includes(next)) {
      throw new InvalidDocumentReviewTransitionError(this.current, next, TRANSITIONS[this.current]);
    }

    if (this.current === next) {
      return;
    }

    console.log(`🔄 [DocumentReviewStateMachine] ${this.current} → ${next}`);
    this.current = next;
  }
}
