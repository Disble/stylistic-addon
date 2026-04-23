import { beforeEach, describe, expect, it, vi } from "vitest";
import type { IFeedbackPort, IDocumentPort } from "../ports";
import type { Suggestion, SuggestionActionResult } from "../types";
import { SuggestionResolutionWorkflow } from "./SuggestionResolutionWorkflow";

/** Builds a canonical resolution suggestion for atomicity-focused workflow tests. */
function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  return {
    id: "s-atomic",
    context: "Contexto con texto original.",
    anchor: "texto original",
    suggestedText: "texto sugerido",
    justification: "Aclara la frase",
    category: "Claridad",
    severity: "medium",
    type: "track-change",
    ...overrides,
  };
}

/** Creates one action result focused on atomic terminal semantics. */
function makeActionResult(
  overrides: Partial<SuggestionActionResult> = {},
): SuggestionActionResult {
  return {
    status: "accepted",
    trackedChangesAffected: 2,
    commentDeleted: true,
    pendingAfter: {
      pendingStylisticArtifacts: 0,
      hasPendingStylisticArtifacts: false,
      trackChangesActive: true,
    },
    documentState: "ready-to-disable-track-changes",
    ...overrides,
  };
}

describe("SuggestionResolutionWorkflow atomicity", () => {
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

  it("dispatches feedback only for atomic accepted results", async () => {
    vi.mocked(documentPort.acceptSuggestion).mockResolvedValue(
      makeActionResult({
        status: "accepted",
        trackedChangesAffected: 2,
        commentDeleted: true,
      }),
    );

    const result = await workflow.acceptSuggestion(makeSuggestion(), "Buen cambio");

    expect(result.status).toBe("accepted");
    expect(result.feedbackStatus).toBe("sent");
    expect(feedbackPort.sendFeedback).toHaveBeenCalledOnce();
  });

  it("does not dispatch feedback when atomic cleanup fails and the adapter returns error", async () => {
    vi.mocked(documentPort.acceptSuggestion).mockResolvedValue(
      makeActionResult({
        status: "error",
        trackedChangesAffected: 2,
        commentDeleted: false,
        error: "GeneralException",
        executionReport: {
          attempted: 2,
          completed: 2,
          remaining: 0,
          error: "GeneralException",
        },
      }),
    );

    const result = await workflow.acceptSuggestion(makeSuggestion());

    expect(result.status).toBe("error");
    expect(result.feedbackStatus).toBe("skipped");
    expect(feedbackPort.sendFeedback).not.toHaveBeenCalled();
  });

  it("does not dispatch feedback when atomic reject leaves the workflow in error", async () => {
    vi.mocked(documentPort.rejectSuggestion).mockResolvedValue(
      makeActionResult({
        status: "error",
        trackedChangesAffected: 2,
        commentDeleted: false,
        error: "ItemNotFound",
        executionReport: {
          attempted: 2,
          completed: 2,
          remaining: 0,
          error: "ItemNotFound",
        },
      }),
    );

    const result = await workflow.rejectSuggestion(makeSuggestion());

    expect(result.status).toBe("error");
    expect(result.feedbackStatus).toBe("skipped");
    expect(feedbackPort.sendFeedback).not.toHaveBeenCalled();
  });
});
