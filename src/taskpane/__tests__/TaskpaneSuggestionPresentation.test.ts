import * as React from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import type { ResultsPanelDeps } from "../SuggestionCardRenderer.types";
import { buildResultsSummary } from "../SuggestionCardRenderer";
import {
  getResultsPanelState,
  resetResultsPanelState,
  setResultsPanelData,
} from "../ResultsPanelStore";
import { resetTaskpaneShellState } from "../TaskpaneShellStore";
import { ResultSuggestionCard } from "../components/ResultSuggestionCard";
import { makeSuggestion } from "../TaskpaneTestHelper";

function createResultsPanelDeps(): ResultsPanelDeps {
  return {
    navigateToText: vi.fn().mockResolvedValue({ status: "navigated" }),
    acceptSuggestion: vi.fn().mockResolvedValue({ status: "accepted" }),
    rejectSuggestion: vi.fn().mockResolvedValue({ status: "rejected" }),
  } as unknown as ResultsPanelDeps;
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

describe("taskpane suggestion presentation", () => {
  beforeEach(() => {
    resetResultsPanelState();
    resetTaskpaneShellState();
  });

  afterEach(() => {
    resetResultsPanelState();
    resetTaskpaneShellState();
  });

  it("builds a live-friendly summary with resolved and remaining counts", () => {
    const firstSuggestion = makeSuggestion({ id: "s-1" });
    const secondSuggestion = makeSuggestion({ id: "s-2" });

    expect(
      buildResultsSummary(
        [firstSuggestion, secondSuggestion],
        {
          successCount: 1,
          failedSuggestions: [
            {
              suggestion: secondSuggestion,
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
          trackChangesActivatedForBatch: true,
        },
        [],
        true
      )
    ).toBe(
      "Sobre selección — Te faltan 1 de 1 sugerencia aplicada por revisar. Todavía no resolviste ninguna. 1 no encontrada(s) en el texto."
    );
  });

  it("renders comment-only suggestions without diff blocks and with text action labels", () => {
    const suggestion = makeSuggestion({
      id: "s-co",
      type: "comment-only",
      suggestedText: undefined,
    });

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
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createResultsPanelDeps()
    );

    const markup = renderCardMarkup("s-co");

    expect(markup).not.toContain('data-testid="card-diff"');
    expect(markup).toContain("Entendido");
    expect(markup).toContain("Ignorar");
    expect(markup).toContain('data-testid="card-comment-badge"');
    expect(markup).toContain("comentario");
  });

  it("renders track-change suggestions with diff blocks and symbolic action labels", () => {
    const suggestion = makeSuggestion({
      id: "s-tc",
      type: "track-change",
      suggestedText: "texto sugerido",
    });

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
        trackChangesActivatedForBatch: false,
      },
      [],
      false,
      createResultsPanelDeps()
    );

    const markup = renderCardMarkup("s-tc");

    expect(markup).toContain('data-testid="card-diff"');
    expect(markup).toContain('aria-label="Aceptar sugerencia"');
    expect(markup).toContain('aria-label="Rechazar sugerencia"');
    expect(markup).toContain("Aceptar");
    expect(markup).toContain("Rechazar");
  });

  it('keeps "No encontrado" cards after actionable suggestions in store ordering', () => {
    const firstSuggestion = makeSuggestion({ id: "s-1", anchor: "primero" });
    const missingSuggestion = makeSuggestion({ id: "s-missing", anchor: "faltante" });
    const secondSuggestion = makeSuggestion({ id: "s-2", anchor: "segundo" });

    setResultsPanelData(
      [firstSuggestion, missingSuggestion, secondSuggestion],
      {
        successCount: 2,
        failedSuggestions: [
          {
            suggestion: missingSuggestion,
            reason: "not-found",
            message: "Anchor no encontrado en el contexto",
          },
        ],
        pendingAfter: {
          pendingStylisticArtifacts: 2,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        trackChangesActivatedForBatch: true,
      },
      [],
      false,
      createResultsPanelDeps()
    );

    expect(getResultsPanelState().cards.map((card) => card.suggestion.anchor)).toEqual([
      "primero",
      "segundo",
      "faltante",
    ]);
    expect(renderCardMarkup("s-missing")).toContain("No encontrado: &quot;faltante&quot;");
  });
});
