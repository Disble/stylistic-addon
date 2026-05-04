import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import type { ResultsPanelDeps } from "./SuggestionCardRenderer.types";
import {
  acceptResultsPanelSuggestion,
  getResultsPanelState,
  rejectResultsPanelSuggestion,
  resetResultsPanelState,
  setResultsPanelData,
  setResultsPanelFeedbackComment,
  toggleResultsPanelFeedback,
} from "./ResultsPanelStore";
import { resetTaskpaneShellState } from "./TaskpaneShellStore";
import { ResultSuggestionCard } from "./components/ResultSuggestionCard";
import { makeSuggestion } from "./TaskpaneTestHelper";

function createResultsPanelDeps(overrides: Partial<ResultsPanelDeps> = {}): ResultsPanelDeps {
  return {
    navigateToText: vi.fn().mockResolvedValue({ status: "navigated" }),
    acceptSuggestion: vi.fn().mockResolvedValue({
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
    }),
    rejectSuggestion: vi.fn().mockResolvedValue({
      status: "rejected",
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
    }),
    ...overrides,
  };
}

function seedResultsPanel(deps: ResultsPanelDeps, failed = false) {
  const suggestion = makeSuggestion({ id: failed ? "s-fail" : "s-1" });
  setResultsPanelData(
    [suggestion],
    {
      successCount: failed ? 0 : 1,
      failedSuggestions: failed
        ? [
            {
              suggestion,
              reason: "not-found",
              message: "Anchor no encontrado en el contexto",
            },
          ]
        : [],
      pendingAfter: {
        pendingStylisticArtifacts: failed ? 0 : 1,
        hasPendingStylisticArtifacts: !failed,
        trackChangesActive: !failed,
      },
      documentState: failed ? "idle" : "pending-review",
      trackChangesActivatedForBatch: !failed,
    },
    [],
    false,
    deps
  );
  return suggestion;
}

function renderCardMarkup(cardId: string): string {
  const card = getResultsPanelState().cards.find((entry) => entry.suggestion.id === cardId);
  if (!card) {
    throw new Error(`Missing card: ${cardId}`);
  }

  return renderToStaticMarkup(
    React.createElement(ResultSuggestionCard, {
      card,
      onAccept: async () => {},
      onFeedbackCommentChange: () => {},
      onNavigate: async () => {},
      onReject: async () => {},
      onToggleFeedback: () => {},
    })
  );
}

describe("taskpane feedback controls", () => {
  beforeEach(() => {
    resetResultsPanelState();
    resetTaskpaneShellState();
  });

  afterEach(() => {
    resetResultsPanelState();
    resetTaskpaneShellState();
  });

  it("renders a feedback button and accordion for each non-failed suggestion", () => {
    seedResultsPanel(createResultsPanelDeps());

    const markup = renderCardMarkup("s-1");

    expect(markup).toContain('data-action="feedback"');
    expect(markup).toContain("feedback-accordion");
    expect(markup).toContain("feedback-textarea");
  });

  it("omits feedback controls for failed suggestions", () => {
    seedResultsPanel(createResultsPanelDeps(), true);

    const markup = renderCardMarkup("s-fail");

    expect(markup).not.toContain('data-action="feedback"');
    expect(markup).not.toContain("feedback-accordion");
  });

  it("toggles the feedback accordion state when requested", () => {
    seedResultsPanel(createResultsPanelDeps());

    expect(getResultsPanelState().cards[0].feedbackOpen).toBe(false);
    toggleResultsPanelFeedback("s-1");
    expect(getResultsPanelState().cards[0].feedbackOpen).toBe(true);
    toggleResultsPanelFeedback("s-1");
    expect(getResultsPanelState().cards[0].feedbackOpen).toBe(false);
  });

  it("omits empty textarea comments from accept resolution payloads", async () => {
    const deps = createResultsPanelDeps();
    const suggestion = seedResultsPanel(deps);

    setResultsPanelFeedbackComment(suggestion.id, "   ");
    await acceptResultsPanelSuggestion(suggestion.id);

    expect(deps.acceptSuggestion).toHaveBeenCalledWith(suggestion, undefined);
  });

  it("includes non-empty textarea comments in accept resolution payloads", async () => {
    const deps = createResultsPanelDeps();
    const suggestion = seedResultsPanel(deps);

    setResultsPanelFeedbackComment(suggestion.id, "Muy buen cambio");
    await acceptResultsPanelSuggestion(suggestion.id);

    expect(deps.acceptSuggestion).toHaveBeenCalledWith(suggestion, "Muy buen cambio");
  });

  it("includes non-empty textarea comments in reject resolution payloads", async () => {
    const deps = createResultsPanelDeps();
    const suggestion = seedResultsPanel(deps);

    setResultsPanelFeedbackComment(suggestion.id, "No me convence");
    await rejectResultsPanelSuggestion(suggestion.id);

    expect(deps.rejectSuggestion).toHaveBeenCalledWith(suggestion, "No me convence");
  });
});
