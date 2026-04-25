import type {
  DocumentReviewState,
  ReplaceSuggestionIdentity,
  SuggestionObservationStatus,
} from "../../../domain/types";
import type { ReplaceTrackedChangeSide } from "./ReplaceResolutionStrategyContext";

/** Terminal success states produced by resolution workflows. */
export type ResolutionStatus =
  | "accepted"
  | "rejected"
  | "unobservable"
  | "identity-lost"
  | "ambiguous-location"
  | "mixed-group";

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
  debugMetadata?: ResolutionObservationDebugMetadata;
  semanticCandidates?: ReplaceSemanticCandidateMap;
  group?: OperationalWrapperGroup;
};

/** One wrapper selected as part of an explicit operational replace group. */
export type OperationalWrapperGroupMember = {
  cc: Word.ContentControl;
  identity: ReplaceSuggestionIdentity;
};

/** Explicit contiguous replace group resolved from operational wrapper metadata. */
export type OperationalWrapperGroup = {
  groupId: string;
  members: OperationalWrapperGroupMember[];
  status: "single" | "contiguous" | "mixed" | "ambiguous";
};

/** One executable tracked-change candidate together with the evidence source that exposed it. */
export type ReplaceTrackedChangeCandidate = {
  trackedChange: Word.TrackedChange;
  source: string;
};

/** Exhaustive tracked-change candidates grouped by semantic side. */
export type ReplaceSemanticCandidateMap = Record<
  ReplaceTrackedChangeSide,
  ReplaceTrackedChangeCandidate[]
>;

/** Primitive debug metadata captured during observation for telemetry/logging. */
export type ResolutionObservationDebugMetadata = {
  selectedCcTag?: string;
  selectedCcTitleKind?: string;
  selectedCommentFound?: boolean;
  observationStatus?: SuggestionObservationStatus;
  identityVersion?: string;
  ccTrackedChangesCount?: number;
  ccRangeTrackedChangesCount?: number;
  operationalAnchorTrackedChangesCount?: number;
  operationalAnchorFound?: boolean;
  commentTrackedChangesCount?: number;
  wrapperGroupId?: string;
  wrapperGroupSize?: number;
  wrapperGroupStatus?: string;
};

/** Normalized resolution observation returned to the command orchestrator. */
export type ResolutionObservation = {
  selectedCc: Word.ContentControl;
  selectedComment: ColocatedCommentContext | null;
  trackedChanges: Word.TrackedChange[];
  observationStatus: SuggestionObservationStatus;
  debugMetadata?: ResolutionObservationDebugMetadata;
  semanticCandidates?: ReplaceSemanticCandidateMap;
  group?: OperationalWrapperGroup;
};

/** Result of locating candidate artifacts for one suggestion. */
export type LocatedSuggestionArtifacts = {
  candidates: Word.ContentControl[];
  selectedCc: Word.ContentControl | null;
  locateStatus: SuggestionObservationStatus | "cc-not-found";
};

/** Input for comment-only resolution after the anchor CC is known. */
export type CommentOnlyResolutionRequest = {
  context: Word.RequestContext;
  cc: Word.ContentControl;
  commentDeleted: boolean;
  pendingBefore: DocumentReviewState;
};
