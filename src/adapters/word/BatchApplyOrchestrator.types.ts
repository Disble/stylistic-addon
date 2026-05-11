import type { CommandResult } from "../../domain/DocumentApplication.types";
import type {
  DocumentReviewState,
  DocumentReviewUiState,
} from "../../domain/review/DocumentReviewStateMachine.types";
import type { Suggestion } from "../../domain/suggestion/Suggestion.types";

/** Injected capabilities required by the orchestrator. */
export type BatchApplyDependencies = {
  /** Enables Track Changes lazily; returns `true` when newly activated. */
  ensureTrackChangesActive: () => Promise<boolean>;
  /** Returns the current document-derived review state. */
  getDocumentReviewState: () => Promise<DocumentReviewState>;
  /** Derives the explicit UI state from a review snapshot. */
  deriveDocumentState: (state: DocumentReviewState) => DocumentReviewUiState;
  /** Optionally refreshes one snapshot hint when local patch reseed is insufficient. */
  rereadSuggestionPositionHint?: (
    suggestion: Suggestion,
    patch: NonNullable<CommandResult["mutationPatch"]>
  ) => Promise<Suggestion["positionHint"] | undefined>;
};
