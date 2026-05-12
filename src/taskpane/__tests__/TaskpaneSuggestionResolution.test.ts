import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { SuggestionResolutionMediatorResult } from "../../domain/suggestion/SuggestionResolutionWorkflow.types";
import type { ResultsPanelDeps } from "../SuggestionCardRenderer.types";
import {
  acceptResultsPanelSuggestion,
  getResultsPanelState,
  rejectResultsPanelSuggestion,
  resetResultsPanelState,
  setResultsPanelData,
} from "../ResultsPanelStore";
import { getTaskpaneShellState, resetTaskpaneShellState } from "../TaskpaneShellStore";
import { makeSuggestion } from "./TaskpaneTestHelper";

function makeMediatorResult(
  overrides: Partial<SuggestionResolutionMediatorResult> = {}
): SuggestionResolutionMediatorResult {
  return {
    status: "accepted",
    trackedChangesAffected: 2,
    commentDeleted: true,
    pendingAfter: {
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    },
    documentState: "pending-review",
    feedbackStatus: "sent",
    taskpaneState: {
      documentState: "pending-review",
      showDisableTrackChangesCta: false,
      showCleanupSection: false,
    },
    ...overrides,
  };
}

function createResultsPanelDeps(overrides: Partial<ResultsPanelDeps> = {}): ResultsPanelDeps {
  return {
    navigateToText: vi.fn().mockResolvedValue({ status: "navigated" }),
    acceptSuggestion: vi.fn().mockResolvedValue(makeMediatorResult()),
    rejectSuggestion: vi.fn().mockResolvedValue(
      makeMediatorResult({
        status: "rejected",
      })
    ),
    ...overrides,
  };
}

function seedResultsPanel(deps: ResultsPanelDeps) {
  const suggestion = makeSuggestion({ id: "s-1" });
  setResultsPanelData(
    [suggestion],
    {
      successCount: 1,
      failedSuggestions: [],
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      trackChangesActivatedForBatch: true,
    },
    [],
    false,
    deps
  );
  return suggestion;
}

describe("TaskpaneSuggestionResolution", () => {
  beforeEach(() => {
    resetResultsPanelState();
    resetTaskpaneShellState();
  });

  afterEach(() => {
    resetResultsPanelState();
    resetTaskpaneShellState();
  });

  it("marks an accepted suggestion as terminal and updates taskpane CTAs", async () => {
    const deps = createResultsPanelDeps({
      acceptSuggestion: vi.fn().mockResolvedValue(
        makeMediatorResult({
          status: "accepted",
          pendingAfter: {
            pendingStylisticArtifacts: 0,
            hasPendingStylisticArtifacts: false,
            trackChangesActive: true,
          },
          documentState: "ready-to-disable-track-changes",
          taskpaneState: {
            documentState: "ready-to-disable-track-changes",
            showDisableTrackChangesCta: true,
            showCleanupSection: true,
          },
        })
      ),
    });
    const suggestion = seedResultsPanel(deps);

    await acceptResultsPanelSuggestion(suggestion.id);

    const card = getResultsPanelState().cards[0];
    expect(deps.acceptSuggestion).toHaveBeenCalledWith(suggestion, undefined);
    expect(card.state).toBe("accepted");
    expect(card.cardGroup).toBe("processed");
    expect(card.hideActions).toBe(true);
    expect(getTaskpaneShellState().cleanupVisible).toBe(true);
    expect(getTaskpaneShellState().disableTrackChangesCtaVisible).toBe(true);
  });

  it("marks a rejected suggestion as terminal and removes action buttons", async () => {
    const deps = createResultsPanelDeps();
    const suggestion = seedResultsPanel(deps);

    await rejectResultsPanelSuggestion(suggestion.id);

    const card = getResultsPanelState().cards[0];
    expect(deps.rejectSuggestion).toHaveBeenCalledWith(suggestion, undefined);
    expect(card.state).toBe("rejected");
    expect(card.cardGroup).toBe("processed");
    expect(card.hideActions).toBe(true);
  });

  it("re-enables the card and keeps it non-terminal on unobservable", async () => {
    const deps = createResultsPanelDeps({
      acceptSuggestion: vi.fn().mockResolvedValue(
        makeMediatorResult({
          status: "unobservable",
          trackedChangesAffected: 0,
          commentDeleted: false,
          feedbackStatus: "skipped",
          error: "Word no expuso suficiente evidencia operacional para confirmar la resolución.",
        })
      ),
    });
    const suggestion = seedResultsPanel(deps);

    await acceptResultsPanelSuggestion(suggestion.id);

    const card = getResultsPanelState().cards[0];
    expect(card.hideActions).toBe(false);
    expect(card.isResolving).toBe(false);
    expect(card.cardGroup).toBe("active");
    expect(card.state).toBe("unobservable");
    expect(getTaskpaneShellState().status.message).toBe(
      "Word no expuso suficiente evidencia operacional para confirmar la resolución."
    );
  });

  it("renders ambiguous-location as terminal manual-review state", async () => {
    const deps = createResultsPanelDeps({
      acceptSuggestion: vi.fn().mockResolvedValue(
        makeMediatorResult({
          status: "ambiguous-location",
          trackedChangesAffected: 0,
          commentDeleted: false,
          feedbackStatus: "skipped",
          error: "La ubicación de la sugerencia es ambigua.",
        })
      ),
    });
    const suggestion = seedResultsPanel(deps);

    await acceptResultsPanelSuggestion(suggestion.id);

    const card = getResultsPanelState().cards[0];
    expect(card.state).toBe("ambiguous-location");
    expect(card.hideActions).toBe(true);
    expect(card.resolutionNote).toBe("(resolución ambigua; reanalizá la sugerencia)");
    expect(getTaskpaneShellState().status.message).toBe(
      "La ubicación de la sugerencia es ambigua."
    );
  });

  it("re-enables the card and surfaces an error when mediator resolution rejects", async () => {
    const deps = createResultsPanelDeps({
      acceptSuggestion: vi.fn().mockRejectedValue(new Error("cleanup preview failed")),
    });
    const suggestion = seedResultsPanel(deps);

    await expect(acceptResultsPanelSuggestion(suggestion.id)).resolves.toBeUndefined();

    const card = getResultsPanelState().cards[0];
    expect(card.state).toBe("error");
    expect(card.cardGroup).toBe("active");
    expect(card.hideActions).toBe(false);
    expect(card.isResolving).toBe(false);
    expect(getTaskpaneShellState().status.message).toBe("cleanup preview failed");
  });
});
