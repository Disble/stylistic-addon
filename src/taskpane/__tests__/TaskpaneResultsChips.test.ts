import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import { ResultsSummaryChips } from "../components/ResultsSummaryChips";
import type { ResultsPanelDeps } from "../SuggestionCardRenderer.types";
import {
  acceptResultsPanelSuggestion,
  getResultsPanelState,
  rejectResultsPanelSuggestion,
  resetResultsPanelState,
  setResultsPanelData,
  setResultsPanelFilter,
} from "../ResultsPanelStore";
import {
  computeResultsPanelChipCounts,
  selectResultsPanelVisibleCards,
} from "../ResultsPanelFilters";
import { resetTaskpaneShellState } from "../TaskpaneShellStore";
import { makeSuggestion } from "./TaskpaneTestHelper";
import type { Suggestion, SuggestionSeverity } from "../../domain/suggestion/Suggestion.types";

function createResultsPanelDeps(): ResultsPanelDeps {
  return {
    navigateToText: vi.fn().mockResolvedValue({ status: "navigated" }),
    acceptSuggestion: vi.fn().mockResolvedValue({
      status: "accepted",
      trackedChangesAffected: 1,
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
      trackedChangesAffected: 1,
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
  };
}

function suggestionWithSeverity(id: string, severity: SuggestionSeverity): Suggestion {
  return makeSuggestion({ id, severity, anchor: `texto-${id}` });
}

function seedMixedPanel(deps: ResultsPanelDeps) {
  const high1 = suggestionWithSeverity("h1", "high");
  const high2 = suggestionWithSeverity("h2", "high");
  const medium1 = suggestionWithSeverity("m1", "medium");
  const low1 = suggestionWithSeverity("l1", "low");
  const failedSuggestion = suggestionWithSeverity("nf", "medium");

  setResultsPanelData(
    [high1, high2, medium1, low1, failedSuggestion],
    {
      successCount: 4,
      failedSuggestions: [
        {
          suggestion: failedSuggestion,
          reason: "not-found",
          message: "no encontrada",
        },
      ],
      pendingAfter: {
        pendingStylisticArtifacts: 4,
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
}

describe("results panel chip filter", () => {
  beforeEach(() => {
    resetResultsPanelState();
    resetTaskpaneShellState();
  });

  afterEach(() => {
    resetResultsPanelState();
    resetTaskpaneShellState();
  });

  it("buckets each pending card by its suggestion severity", () => {
    seedMixedPanel(createResultsPanelDeps());

    const counts = computeResultsPanelChipCounts(getResultsPanelState().cards);

    expect(counts).toEqual({
      all: 5,
      high: 2,
      medium: 1,
      low: 1,
      accepted: 0,
      rejected: 0,
      failed: 1,
    });
  });

  it("moves resolved cards out of severity buckets and into the matching state bucket", async () => {
    const deps = createResultsPanelDeps();
    seedMixedPanel(deps);

    await acceptResultsPanelSuggestion("h1");
    await rejectResultsPanelSuggestion("m1");

    const counts = computeResultsPanelChipCounts(getResultsPanelState().cards);

    expect(counts.high).toBe(1);
    expect(counts.medium).toBe(0);
    expect(counts.accepted).toBe(1);
    expect(counts.rejected).toBe(1);
    expect(counts.failed).toBe(1);
  });

  it("filters visible cards to the active chip selection", () => {
    seedMixedPanel(createResultsPanelDeps());
    const allCards = getResultsPanelState().cards;

    const onlyHigh = selectResultsPanelVisibleCards(allCards, "high");
    expect(onlyHigh.map((card) => card.suggestion.id)).toEqual(["h1", "h2"]);

    const onlyFailed = selectResultsPanelVisibleCards(allCards, "failed");
    expect(onlyFailed.map((card) => card.suggestion.id)).toEqual(["nf"]);

    const allFilter = selectResultsPanelVisibleCards(allCards, "all");
    expect(allFilter).toHaveLength(5);
  });

  it("resets the active filter back to all when new pipeline data lands", () => {
    seedMixedPanel(createResultsPanelDeps());
    setResultsPanelFilter("high");
    expect(getResultsPanelState().activeFilter).toBe("high");

    seedMixedPanel(createResultsPanelDeps());
    expect(getResultsPanelState().activeFilter).toBe("all");
  });

  it("renders only buckets with non-zero counts and marks the active chip pressed", () => {
    const counts = {
      all: 4,
      high: 2,
      medium: 0,
      low: 1,
      accepted: 1,
      rejected: 0,
      failed: 0,
    };

    const markup = renderToStaticMarkup(
      React.createElement(ResultsSummaryChips, {
        activeFilter: "high",
        counts,
        onFilterChange: () => {},
        summaryText: "Resumen accesible",
      })
    );

    expect(markup).toContain('data-testid="results-chip-all"');
    expect(markup).toContain('data-testid="results-chip-high"');
    expect(markup).toContain('data-testid="results-chip-low"');
    expect(markup).toContain('data-testid="results-chip-accepted"');
    expect(markup).not.toContain('data-testid="results-chip-medium"');
    expect(markup).not.toContain('data-testid="results-chip-rejected"');
    expect(markup).not.toContain('data-testid="results-chip-failed"');
    expect(markup).toContain('data-active="true"');
    expect(markup).toContain('aria-pressed="true"');
  });
});
