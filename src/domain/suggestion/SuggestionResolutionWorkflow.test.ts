import { beforeEach, describe, expect, it, vi } from "vitest";
import type {
  IFeedbackPort,
  IDocumentPort,
} from "../ports";
import type { Suggestion, SuggestionActionResult } from "../types";
import { SuggestionResolutionWorkflow } from "./SuggestionResolutionWorkflow";

function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  return {
    id: "s-1",
    context: "Contexto con texto original.",
    anchor: "texto original",
    suggestedText: "texto sugerido",
    justification: "Mas claro",
    category: "Claridad",
    severity: "medium",
    type: "track-change",
    ...overrides,
  };
}

function makeActionResult(
  overrides: Partial<SuggestionActionResult> = {},
): SuggestionActionResult {
  return {
    status: "accepted",
    trackedChangesAffected: 1,
    commentDeleted: true,
    pendingAfter: {
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    },
    documentState: "pending-review",
    ...overrides,
  };
}

describe("SuggestionResolutionWorkflow", () => {
  let documentPort: IDocumentPort;
  let feedbackPort: IFeedbackPort;
  let workflow: SuggestionResolutionWorkflow;

  beforeEach(() => {
    documentPort = {
      getTextToAnalyze: vi.fn(),
      getAppliedOriginalTexts: vi.fn(),
      applySuggestions: vi.fn(),
      getCleanupPreview: vi.fn(),
      cleanupResolvedComments: vi.fn(),
      acceptSuggestion: vi.fn(),
      rejectSuggestion: vi.fn(),
      getDocumentReviewState: vi.fn(),
      disableTrackChanges: vi.fn(),
      navigateToText: vi.fn(),
    } as unknown as IDocumentPort;

    feedbackPort = {
      sendFeedback: vi.fn().mockResolvedValue(undefined),
    };

    workflow = new SuggestionResolutionWorkflow(documentPort, feedbackPort);
  });

  it("returns sent feedback status for accepted suggestions and dispatches positive feedback", async () => {
    const suggestion = makeSuggestion();
    vi.mocked(documentPort.acceptSuggestion).mockResolvedValue(
      makeActionResult({ status: "accepted" }),
    );

    const result = await workflow.acceptSuggestion(suggestion, "Muy buen cambio");

    expect(result.feedbackStatus).toBe("sent");
    expect(feedbackPort.sendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({
        originalText: suggestion.anchor,
        rating: "positive",
        comment: "Muy buen cambio",
      }),
    );
  });

  it("returns sent feedback status for rejected suggestions and dispatches negative feedback", async () => {
    const suggestion = makeSuggestion();
    vi.mocked(documentPort.rejectSuggestion).mockResolvedValue(
      makeActionResult({ status: "rejected" }),
    );

    const result = await workflow.rejectSuggestion(suggestion);

    expect(result.feedbackStatus).toBe("sent");
    expect(feedbackPort.sendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({
        originalText: suggestion.anchor,
        rating: "negative",
      }),
    );
  });

  it("skips feedback for cc-not-found terminal results", async () => {
    vi.mocked(documentPort.acceptSuggestion).mockResolvedValue(
      makeActionResult({
        status: "cc-not-found",
        trackedChangesAffected: 0,
        commentDeleted: false,
      }),
    );

    const result = await workflow.acceptSuggestion(makeSuggestion());

    expect(result.feedbackStatus).toBe("skipped");
    expect(feedbackPort.sendFeedback).not.toHaveBeenCalled();
  });

  it("skips feedback for generic error results", async () => {
    vi.mocked(documentPort.rejectSuggestion).mockResolvedValue(
      makeActionResult({
        status: "error",
        trackedChangesAffected: 0,
        commentDeleted: false,
        error: "El documento está protegido",
      }),
    );

    const result = await workflow.rejectSuggestion(makeSuggestion());

    expect(result.feedbackStatus).toBe("skipped");
    expect(feedbackPort.sendFeedback).not.toHaveBeenCalled();
  });

  it("skips feedback for unobservable results", async () => {
    vi.mocked(documentPort.acceptSuggestion).mockResolvedValue(
      makeActionResult({
        status: "unobservable",
        trackedChangesAffected: 0,
        commentDeleted: false,
        error:
          "Word no expuso suficientes tracked changes para confirmar la resolución.",
      }),
    );

    const result = await workflow.acceptSuggestion(makeSuggestion());

    expect(result.feedbackStatus).toBe("skipped");
    expect(feedbackPort.sendFeedback).not.toHaveBeenCalled();
  });

  it("returns failed feedback status when the feedback port throws synchronously", async () => {
    vi.mocked(documentPort.acceptSuggestion).mockResolvedValue(
      makeActionResult({ status: "accepted" }),
    );
    feedbackPort = {
      sendFeedback: vi.fn(() => {
        throw new Error("sync failure");
      }),
    };
    workflow = new SuggestionResolutionWorkflow(documentPort, feedbackPort);

    const result = await workflow.acceptSuggestion(makeSuggestion());

    expect(result.feedbackStatus).toBe("failed");
  });
});
