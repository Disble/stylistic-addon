import type {
  DocumentReviewState,
  ReplaceSuggestionIdentity,
  SuggestionObservationStatus,
} from "../../../domain/types";

/** Terminal success states produced by resolution workflows. */
export type ResolutionStatus =
  | "accepted"
  | "rejected"
  | "unobservable"
  | "identity-lost";

/** Stylistic comment plus its actionable range. */
export type ColocatedCommentContext = {
  comment: Word.Comment;
  range: Word.Range;
};

/** Replace-specific evidence gathered while observing a suggestion. */
export type ReplaceObservationContext = {
  identity?: ReplaceSuggestionIdentity;
  trackedChanges: Word.TrackedChange[];
  observationStatus: SuggestionObservationStatus;
};

/** Normalized resolution observation returned to the command orchestrator. */
export type ResolutionObservation = {
  selectedCc: Word.ContentControl;
  selectedComment: ColocatedCommentContext | null;
  trackedChanges: Word.TrackedChange[];
  observationStatus: SuggestionObservationStatus;
};

/** Result of locating candidate artifacts for one suggestion. */
export type LocatedSuggestionArtifacts = {
  rankedCandidates: Word.ContentControl[];
  selectedCc: Word.ContentControl | null;
};

/** Input for comment-only resolution after the anchor CC is known. */
export type CommentOnlyResolutionRequest = {
  context: Word.RequestContext;
  cc: Word.ContentControl;
  commentDeleted: boolean;
  pendingBefore: DocumentReviewState;
};
