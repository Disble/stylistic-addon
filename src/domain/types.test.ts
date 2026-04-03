import type { IDocumentPort } from "./ports";
import type {
  ApplySuggestionsResult,
  DocumentReviewState,
  Suggestion,
  SuggestionActionResult,
  SuggestionResolutionWorkflowResult,
  SuggestionState,
  WorkflowSuggestion,
} from "./types";

// ---------------------------------------------------------------------------
// SuggestionState — compile-time assignability checks
// ---------------------------------------------------------------------------

const _state1: SuggestionState = "pending";
const _state2: SuggestionState = "accepted";
const _state3: SuggestionState = "rejected";
const _state4: SuggestionState = "already-resolved";

// Suppress unused variable warnings
void _state1;
void _state2;
void _state3;
void _state4;

// ---------------------------------------------------------------------------
// SuggestionActionResult — shape checks
// ---------------------------------------------------------------------------

const _result: SuggestionActionResult = {
  status: "accepted",
  trackedChangesAffected: 2,
  commentDeleted: true,
  pendingAfter: {
    pendingStylisticArtifacts: 1,
    hasPendingStylisticArtifacts: true,
    trackChangesActive: true,
  },
  transitionedToZeroPending: false,
  showDisableTrackChangesCta: false,
};

const _resultWithError: SuggestionActionResult = {
  status: "error",
  trackedChangesAffected: 0,
  commentDeleted: false,
  pendingAfter: {
    pendingStylisticArtifacts: 0,
    hasPendingStylisticArtifacts: false,
    trackChangesActive: false,
  },
  transitionedToZeroPending: false,
  showDisableTrackChangesCta: false,
  error: "something went wrong",
};

const _reviewState: DocumentReviewState = {
  pendingStylisticArtifacts: 1,
  hasPendingStylisticArtifacts: true,
  trackChangesActive: true,
};

const _applyResult: ApplySuggestionsResult = {
  successCount: 1,
  failedSuggestions: [],
  pendingAfter: _reviewState,
  trackChangesActivatedForBatch: true,
};

const _workflowResult: SuggestionResolutionWorkflowResult = {
  status: "accepted",
  trackedChangesAffected: 1,
  commentDeleted: true,
  pendingAfter: _reviewState,
  transitionedToZeroPending: false,
  showDisableTrackChangesCta: false,
  feedbackStatus: "sent",
};

const _suggestion: Suggestion = {
  id: "s1",
  context: "El texto original aparece dentro de este contexto.",
  anchor: "texto original",
  suggestedText: "texto mejorado",
  justification: "Mejora la claridad",
  category: "Claridad",
  severity: "medium",
  type: "track-change",
};

const _workflowSuggestion: WorkflowSuggestion = {
  context: "El texto original aparece dentro de este contexto.",
  anchor: "texto original",
  suggestedText: "texto mejorado",
  justification: "Mejora la claridad",
  category: "Claridad",
  severity: "medium",
};

void _result;
void _resultWithError;
void _suggestion;
void _workflowSuggestion;
void _reviewState;
void _applyResult;
void _workflowResult;

type _SuggestionHasOriginalText = Suggestion extends { originalText: string }
  ? true
  : false;
type _WorkflowSuggestionHasOriginalText = WorkflowSuggestion extends {
  originalText: string;
}
  ? true
  : false;
const _checkSuggestionOriginalTextRemoved: _SuggestionHasOriginalText = false;
const _checkWorkflowOriginalTextRemoved: _WorkflowSuggestionHasOriginalText =
  false;

void _checkSuggestionOriginalTextRemoved;
void _checkWorkflowOriginalTextRemoved;

// ---------------------------------------------------------------------------
// IDocumentPort — acceptSuggestion and rejectSuggestion method presence
// ---------------------------------------------------------------------------

type _HasAccept = IDocumentPort extends {
  acceptSuggestion(s: Suggestion): Promise<SuggestionActionResult>;
}
  ? true
  : false;
type _HasReject = IDocumentPort extends {
  rejectSuggestion(s: Suggestion): Promise<SuggestionActionResult>;
}
  ? true
  : false;
type _HasReviewState = IDocumentPort extends {
  getDocumentReviewState(): Promise<DocumentReviewState>;
}
  ? true
  : false;
type _HasDisableTrackChanges = IDocumentPort extends {
  disableTrackChanges(): Promise<void>;
}
  ? true
  : false;
const _checkAccept: _HasAccept = true;
const _checkReject: _HasReject = true;
const _checkReviewState: _HasReviewState = true;
const _checkDisableTrackChanges: _HasDisableTrackChanges = true;

void _checkAccept;
void _checkReject;
void _checkReviewState;
void _checkDisableTrackChanges;

// ---------------------------------------------------------------------------
// Runtime no-op test so Vitest registers the file
// ---------------------------------------------------------------------------

import { describe, it } from "vitest";

describe("domain types (compile-time checks)", () => {
  it("passes if the file compiles without errors", () => {
    // All checks are compile-time — if this file compiles, the types are correct.
  });
});

// ---------------------------------------------------------------------------
// FeedbackPayload — compile-time shape checks
// ---------------------------------------------------------------------------

import type { FeedbackPayload } from "./types";

const _feedbackMinimal: FeedbackPayload = {
  category: "Redundancia",
  originalText: "completamente necesario",
  suggestedText: "necesario",
  justification: "Ya implica completitud.",
  rating: "positive",
  severity: "high",
};

const _feedbackWithComment: FeedbackPayload = {
  category: "Muletilla",
  originalText: "básicamente",
  suggestedText: "",
  justification: "Filler word.",
  rating: "negative",
  severity: "medium",
  comment: "Estoy de acuerdo",
};

void _feedbackMinimal;
void _feedbackWithComment;

// FeedbackPayload must have a justification field
type _HasJustification = FeedbackPayload extends { justification: string }
  ? true
  : false;
const _checkJustification: _HasJustification = true;
void _checkJustification;

// rating must be exactly "positive" | "negative"
type _RatingPositive = "positive" extends FeedbackPayload["rating"]
  ? true
  : false;
type _RatingNegative = "negative" extends FeedbackPayload["rating"]
  ? true
  : false;
const _r1: _RatingPositive = true;
const _r2: _RatingNegative = true;
void _r1;
void _r2;

// comment must be optional (can construct without it — already verified by _feedbackMinimal above)
// Just verify comment field IS optional by checking it's not in the Required<> keys contract
type _WithComment = Required<FeedbackPayload>;
const _checkOptional: _WithComment["comment"] = "test";
void _checkOptional;

// ---------------------------------------------------------------------------
// IFeedbackPort — compile-time checks
// ---------------------------------------------------------------------------

import type { IFeedbackPort } from "./ports";

// IFeedbackPort must have sendFeedback(payload: FeedbackPayload): Promise<void>
type _HasSendFeedback = IFeedbackPort extends {
  sendFeedback(payload: FeedbackPayload): Promise<void>;
}
  ? true
  : false;
const _checkSendFeedback: _HasSendFeedback = true;
void _checkSendFeedback;

describe("IFeedbackPort (compile-time checks)", () => {
  it("declares sendFeedback(payload: FeedbackPayload): Promise<void>", () => {
    // Checked above
  });
});
