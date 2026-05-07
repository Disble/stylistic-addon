import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import type { ResultsPanelDeps } from "../SuggestionCardRenderer.types";
import {
  acceptResultsPanelSuggestion,
  getResultsPanelState,
  rejectResultsPanelSuggestion,
  resetResultsPanelState,
  setResultsPanelData,
} from "../ResultsPanelStore";
import { resetTaskpaneShellState } from "../TaskpaneShellStore";
import { ResultSuggestionCard } from "../components/ResultSuggestionCard";
import { makeSuggestion } from "../TaskpaneTestHelper";

function createDeps(overrides?: Partial<ResultsPanelDeps>): ResultsPanelDeps {
  return {
    navigateToText: vi.fn().mockResolvedValue({ status: "navigated" }),
    acceptSuggestion: vi.fn().mockResolvedValue({
      status: "accepted",
      taskpaneState: { showDisableTrackChangesCta: false, showCleanupSection: false },
    }),
    rejectSuggestion: vi.fn().mockResolvedValue({
      status: "rejected",
      taskpaneState: { showDisableTrackChangesCta: false, showCleanupSection: false },
    }),
    ...overrides,
  } as unknown as ResultsPanelDeps;
}

function renderCard(cardId: string): string {
  const card = getResultsPanelState().cards.find((entry) => entry.suggestion.id === cardId);
  if (!card) throw new Error(`Missing card: ${cardId}`);
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

describe("suggestion card visuals", () => {
  beforeEach(() => {
    resetResultsPanelState();
    resetTaskpaneShellState();
  });

  afterEach(() => {
    resetResultsPanelState();
    resetTaskpaneShellState();
  });

  it("renders the category as a custom pill with severity dot indicator (no Fluent badge for category)", () => {
    setResultsPanelData(
      [makeSuggestion({ id: "s-1", category: "gramatica", severity: "high" })],
      {
        successCount: 1,
        failedSuggestions: [],
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createDeps()
    );

    const markup = renderCard("s-1");

    expect(markup).toContain('data-testid="card-category-pill"');
    expect(markup).toContain('data-testid="card-severity-indicator"');
    expect(markup).toContain('aria-label="Severidad alta"');
    expect(markup).toContain('data-category-accent="grammar"');
  });

  it("normalizes raw category strings into stable accent slugs", () => {
    setResultsPanelData(
      [
        makeSuggestion({ id: "s-grammar", category: "gramática" }),
        makeSuggestion({ id: "s-spelling", category: "ortografía" }),
        makeSuggestion({ id: "s-punctuation", category: "puntuacion" }),
        makeSuggestion({ id: "s-style", category: "estilo" }),
        makeSuggestion({ id: "s-other", category: "claridad" }),
      ],
      {
        successCount: 5,
        failedSuggestions: [],
        pendingAfter: {
          pendingStylisticArtifacts: 5,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createDeps()
    );

    expect(renderCard("s-grammar")).toContain('data-category-accent="grammar"');
    expect(renderCard("s-spelling")).toContain('data-category-accent="spelling"');
    expect(renderCard("s-punctuation")).toContain('data-category-accent="punctuation"');
    expect(renderCard("s-style")).toContain('data-category-accent="style"');
    expect(renderCard("s-other")).toContain('data-category-accent="neutral"');
  });

  it("flags pending cards with the pending visual state", () => {
    setResultsPanelData(
      [makeSuggestion({ id: "s-1" })],
      {
        successCount: 1,
        failedSuggestions: [],
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createDeps()
    );

    expect(renderCard("s-1")).toContain('data-card-state="pending"');
  });

  it("transitions the card visual state to accepted after a successful accept", async () => {
    setResultsPanelData(
      [makeSuggestion({ id: "s-1" })],
      {
        successCount: 1,
        failedSuggestions: [],
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createDeps()
    );

    await acceptResultsPanelSuggestion("s-1");

    expect(renderCard("s-1")).toContain('data-card-state="accepted"');
  });

  it("transitions the card visual state to rejected after a reject", async () => {
    setResultsPanelData(
      [makeSuggestion({ id: "s-1" })],
      {
        successCount: 1,
        failedSuggestions: [],
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createDeps()
    );

    await rejectResultsPanelSuggestion("s-1");

    expect(renderCard("s-1")).toContain('data-card-state="rejected"');
  });

  it("does not render a status icon while the card is pending", () => {
    setResultsPanelData(
      [makeSuggestion({ id: "s-1" })],
      {
        successCount: 1,
        failedSuggestions: [],
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createDeps()
    );

    const markup = renderCard("s-1");

    expect(markup).not.toContain('data-testid="card-status-icon-accepted"');
    expect(markup).not.toContain('data-testid="card-status-icon-rejected"');
    expect(markup).not.toContain('data-testid="card-status-icon-failed"');
    expect(markup).not.toContain('data-testid="card-status-icon-not-found"');
  });

  it("renders the accepted status icon after accepting a suggestion", async () => {
    setResultsPanelData(
      [makeSuggestion({ id: "s-1" })],
      {
        successCount: 1,
        failedSuggestions: [],
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createDeps()
    );

    await acceptResultsPanelSuggestion("s-1");

    expect(renderCard("s-1")).toContain('data-testid="card-status-icon-accepted"');
  });

  it("renders the not-found status icon for cards whose anchor was not located", () => {
    const okSuggestion = makeSuggestion({ id: "s-ok" });
    const missingSuggestion = makeSuggestion({ id: "s-missing", anchor: "faltante" });

    setResultsPanelData(
      [okSuggestion, missingSuggestion],
      {
        successCount: 1,
        failedSuggestions: [
          {
            suggestion: missingSuggestion,
            reason: "not-found",
            message: "Anchor no encontrado",
          },
        ],
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createDeps()
    );

    expect(renderCard("s-missing")).toContain('data-testid="card-status-icon-not-found"');
  });

  it("does not render the category dot pattern inside the pill anymore", () => {
    setResultsPanelData(
      [makeSuggestion({ id: "s-1", category: "gramatica" })],
      {
        successCount: 1,
        failedSuggestions: [],
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createDeps()
    );

    const markup = renderCard("s-1");
    const pillSnippet = markup.slice(
      markup.indexOf('data-testid="card-category-pill"'),
      markup.indexOf("</span>", markup.indexOf('data-testid="card-category-pill"')) +
        "</span>".length
    );

    expect(pillSnippet).not.toContain('aria-hidden="true"');
  });

  it("flags failed and not-found cards with their own visual states", () => {
    const okSuggestion = makeSuggestion({ id: "s-ok" });
    const missingSuggestion = makeSuggestion({ id: "s-missing", anchor: "faltante" });

    setResultsPanelData(
      [okSuggestion, missingSuggestion],
      {
        successCount: 1,
        failedSuggestions: [
          {
            suggestion: missingSuggestion,
            reason: "not-found",
            message: "Anchor no encontrado",
          },
        ],
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createDeps()
    );

    expect(renderCard("s-missing")).toContain('data-card-state="not-found"');
  });
});
