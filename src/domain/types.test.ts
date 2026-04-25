import { describe, expectTypeOf, it } from "vitest";
import type { IDocumentPort, IFeedbackPort } from "./ports";
import type {
  ApplySuggestionsResult,
  DocumentReviewState,
  FeedbackPayload,
  ReplaceSuggestionIdentity,
  Suggestion,
  SuggestionActionResult,
  SuggestionObservationStatus,
  SuggestionResolutionWorkflowResult,
  SuggestionState,
  WordArtifactRef,
  WorkflowSuggestion,
} from "./types";
import type { DocumentReviewUiState } from "./review/DocumentReviewStateMachine";

describe("domain types (compile-time checks)", () => {
  it("accepts the supported suggestion states", () => {
    expectTypeOf<SuggestionState>().toEqualTypeOf<
      | "pending"
      | "resolving"
      | "accepted"
      | "rejected"
      | "unobservable"
      | "identity-lost"
      | "ambiguous-location"
      | "mixed-group"
      | "error"
    >();
  });

  it("accepts the supported observation statuses", () => {
    expectTypeOf<SuggestionObservationStatus>().toEqualTypeOf<
      | "confirmed-pending"
      | "confirmed-resolved"
      | "unobservable"
      | "identity-lost"
      | "ambiguous-location"
      | "mixed-group"
    >();
  });

  it("supports the persisted replace suggestion identity shape", () => {
    const artifactRef: WordArtifactRef = {
      kind: "content-control",
      role: "inserted-side",
      value: "stylistic:track-change:s1",
    };

    const replaceIdentity: ReplaceSuggestionIdentity = {
      suggestionId: "s1",
      version: "operational-wrapper-v1",
      insertedSideRef: artifactRef,
      deletedSideRef: {
        kind: "anchor",
        role: "deleted-side",
        value: "texto original",
      },
      anchorRef: {
        kind: "anchor",
        role: "operational-anchor",
        value: "Contexto con texto original.",
      },
      groupId: "s1",
      groupIndex: 0,
      groupSize: 1,
    };

    expectTypeOf(artifactRef).toEqualTypeOf<WordArtifactRef>();
    expectTypeOf(replaceIdentity).toEqualTypeOf<ReplaceSuggestionIdentity>();
  });

  it("preserves the document result contracts", () => {
    const reviewState: DocumentReviewState = {
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    };
    const documentUiState: DocumentReviewUiState =
      "ready-to-disable-track-changes";
    const actionResult: SuggestionActionResult = {
      status: "accepted",
      trackedChangesAffected: 2,
      commentDeleted: true,
      pendingAfter: reviewState,
      documentState: "pending-review",
    };
    const actionErrorResult: SuggestionActionResult = {
      status: "error",
      trackedChangesAffected: 0,
      commentDeleted: false,
      pendingAfter: {
        pendingStylisticArtifacts: 0,
        hasPendingStylisticArtifacts: false,
        trackChangesActive: false,
      },
      documentState: "idle",
      error: "something went wrong",
      errorPhase: "execute",
    };
    const applyResult: ApplySuggestionsResult = {
      successCount: 1,
      failedSuggestions: [],
      pendingAfter: reviewState,
      documentState: "pending-review",
      trackChangesActivatedForBatch: true,
    };
    const workflowResult: SuggestionResolutionWorkflowResult = {
      status: "accepted",
      trackedChangesAffected: 1,
      commentDeleted: true,
      pendingAfter: reviewState,
      documentState: "pending-review",
      feedbackStatus: "sent",
    };

    expectTypeOf(reviewState).toEqualTypeOf<DocumentReviewState>();
    expectTypeOf(documentUiState).toMatchTypeOf<DocumentReviewUiState>();
    expectTypeOf(actionResult).toEqualTypeOf<SuggestionActionResult>();
    expectTypeOf(actionErrorResult).toEqualTypeOf<SuggestionActionResult>();
    expectTypeOf(applyResult).toEqualTypeOf<ApplySuggestionsResult>();
    expectTypeOf(workflowResult).toEqualTypeOf<SuggestionResolutionWorkflowResult>();
  });

  it("keeps suggestion contracts free of originalText", () => {
    type SuggestionHasOriginalText = "originalText" extends keyof Suggestion
      ? true
      : false;
    type WorkflowSuggestionHasOriginalText =
      "originalText" extends keyof WorkflowSuggestion ? true : false;

    expectTypeOf<SuggestionHasOriginalText>().toEqualTypeOf<false>();
    expectTypeOf<WorkflowSuggestionHasOriginalText>().toEqualTypeOf<false>();
  });

  it("preserves the suggestion payload shapes", () => {
    const suggestion: Suggestion = {
      id: "s1",
      context: "El texto original aparece dentro de este contexto.",
      anchor: "texto original",
      suggestedText: "texto mejorado",
      justification: "Mejora la claridad",
      category: "Claridad",
      severity: "medium",
      type: "track-change",
    };
    const workflowSuggestion: WorkflowSuggestion = {
      context: "El texto original aparece dentro de este contexto.",
      anchor: "texto original",
      suggestedText: "texto mejorado",
      justification: "Mejora la claridad",
      category: "Claridad",
      severity: "medium",
    };

    expectTypeOf(suggestion).toEqualTypeOf<Suggestion>();
    expectTypeOf(workflowSuggestion).toEqualTypeOf<WorkflowSuggestion>();
  });

  it("exposes the expected review methods on IDocumentPort", () => {
    expectTypeOf<IDocumentPort["acceptSuggestion"]>().toEqualTypeOf<
      (suggestion: Suggestion) => Promise<SuggestionActionResult>
    >();
    expectTypeOf<IDocumentPort["rejectSuggestion"]>().toEqualTypeOf<
      (suggestion: Suggestion) => Promise<SuggestionActionResult>
    >();
    expectTypeOf<IDocumentPort["getDocumentReviewState"]>().toEqualTypeOf<
      () => Promise<DocumentReviewState>
    >();
    expectTypeOf<IDocumentPort["disableTrackChanges"]>().toEqualTypeOf<
      () => Promise<void>
    >();
  });
});

describe("IFeedbackPort (compile-time checks)", () => {
  it("preserves the feedback payload shape", () => {
    const feedbackMinimal: FeedbackPayload = {
      autorSlug: "disble",
      category: "Redundancia",
      context: "Frase con completamente necesario.",
      anchor: "completamente necesario",
      suggestedText: "necesario",
      justification: "Ya implica completitud.",
      action: "accept",
      severity: "high",
      suggestionType: "track-change",
    };
    const feedbackWithComment: FeedbackPayload = {
      autorSlug: "disble",
      category: "Muletilla",
      context: "Frase con básicamente.",
      anchor: "básicamente",
      suggestedText: "",
      justification: "Filler word.",
      action: "reject",
      severity: "medium",
      suggestionType: "track-change",
      comment: "Estoy de acuerdo",
    };

    expectTypeOf(feedbackMinimal).toEqualTypeOf<FeedbackPayload>();
    expectTypeOf(feedbackWithComment).toEqualTypeOf<FeedbackPayload>();
    expectTypeOf<FeedbackPayload["justification"]>().toEqualTypeOf<string>();
    expectTypeOf<FeedbackPayload["action"]>().toEqualTypeOf<"accept" | "reject">();

    type HasOptionalComment = {} extends Pick<FeedbackPayload, "comment">
      ? true
      : false;

    expectTypeOf<HasOptionalComment>().toEqualTypeOf<true>();
  });

  it("declares sendFeedback(payload: FeedbackPayload): Promise<void>", () => {
    expectTypeOf<IFeedbackPort["sendFeedback"]>().toEqualTypeOf<
      (payload: FeedbackPayload) => Promise<void>
    >();
  });
});
