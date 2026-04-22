import { beforeEach, describe, expect, it, vi } from "vitest";
import type { IDocumentPort, IFeedbackPort } from "../ports";
import { SuggestionResolutionWorkflow } from "../suggestion/SuggestionResolutionWorkflow";
import type { Suggestion, SuggestionActionResult } from "../types";
import { ReviewSessionMediator } from "./ReviewSessionMediator";

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

describe("ReviewSessionMediator", () => {
  let documentPort: IDocumentPort;
  let workflow: SuggestionResolutionWorkflow;
  let mediator: ReviewSessionMediator;

  beforeEach(() => {
    documentPort = {
      getTextToAnalyze: vi.fn(),
      getAppliedOriginalTexts: vi.fn(),
      applySuggestions: vi.fn(),
      getCleanupPreview: vi.fn().mockResolvedValue({ deletable: 0, kept: 1 }),
      cleanupResolvedComments: vi.fn(),
      acceptSuggestion: vi.fn(),
      rejectSuggestion: vi.fn(),
      getDocumentReviewState: vi.fn().mockResolvedValue({
        pendingStylisticArtifacts: 0,
        hasPendingStylisticArtifacts: false,
        trackChangesActive: true,
      }),
      disableTrackChanges: vi.fn().mockResolvedValue(undefined),
      navigateToText: vi.fn(),
    } as unknown as IDocumentPort;

    const feedbackPort: IFeedbackPort = {
      sendFeedback: vi.fn().mockResolvedValue(undefined),
    };

    workflow = new SuggestionResolutionWorkflow(documentPort, feedbackPort);
    mediator = new ReviewSessionMediator(documentPort, feedbackPort);
  });

  it("rehydrates taskpane state from the document snapshot plus cleanup preview", async () => {
    vi.mocked(documentPort.getCleanupPreview).mockResolvedValueOnce({
      deletable: 2,
      kept: 0,
    });

    const state = await mediator.rehydrateTaskpaneState();

    expect(state).toEqual({
      documentState: "ready-to-disable-track-changes",
      showDisableTrackChangesCta: true,
      showCleanupSection: true,
    });
  });

  it("returns centralized taskpane state after accept resolution", async () => {
    vi.mocked(documentPort.acceptSuggestion).mockResolvedValue(
      makeActionResult({
        pendingAfter: {
          pendingStylisticArtifacts: 0,
          hasPendingStylisticArtifacts: false,
          trackChangesActive: true,
        },
        documentState: "ready-to-disable-track-changes",
      }),
    );
    vi.mocked(documentPort.getCleanupPreview).mockResolvedValueOnce({
      deletable: 1,
      kept: 0,
    });

    const result = await mediator.acceptSuggestion(makeSuggestion(), "ok");

    expect(result.taskpaneState).toEqual({
      documentState: "ready-to-disable-track-changes",
      showDisableTrackChangesCta: true,
      showCleanupSection: true,
    });
  });

  it("preserves workflow warnings while enriching taskpane state", async () => {
    vi.mocked(documentPort.acceptSuggestion).mockResolvedValue(
      makeActionResult({
        warnings: [
          {
            code: "cleanup-failed",
            phase: "cleanup",
            message: "late ItemNotFound",
          },
        ],
      }),
    );

    const result = await mediator.acceptSuggestion(makeSuggestion());

    expect(result.warnings).toEqual([
      {
        code: "cleanup-failed",
        phase: "cleanup",
        message: "late ItemNotFound",
      },
    ]);
    expect(result.taskpaneState).toEqual({
      documentState: "pending-review",
      showDisableTrackChangesCta: false,
      showCleanupSection: false,
    });
  });

  it("delegates explicit Track Changes deactivation and returns updated taskpane state", async () => {
    vi.mocked(documentPort.getCleanupPreview).mockResolvedValueOnce({
      deletable: 0,
      kept: 0,
    });
    await mediator.rehydrateTaskpaneState();
    vi.mocked(documentPort.getCleanupPreview).mockResolvedValueOnce({
      deletable: 0,
      kept: 0,
    });

    const state = await mediator.disableTrackChanges();

    expect(documentPort.disableTrackChanges).toHaveBeenCalledOnce();
    expect(state).toEqual({
      documentState: "idle",
      showDisableTrackChangesCta: false,
      showCleanupSection: false,
    });
  });
});
