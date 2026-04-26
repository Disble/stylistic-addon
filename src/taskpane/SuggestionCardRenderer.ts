/* global document */

/**
 * SuggestionCardRenderer — taskpane-facing facade for suggestion-card rendering.
 *
 * The concrete DOM, action, state, feedback, and list responsibilities live in
 * focused `suggestion-card/*` modules. This file owns only the public API used
 * by the taskpane composition root.
 *
 * @module SuggestionCardRenderer
 */

import type { ApplySuggestionsResult } from "../domain/DocumentApplication.types";
import type { Suggestion } from "../domain/suggestion/Suggestion.types";
import type { ResultsPanelDeps } from "./SuggestionCardRenderer.types";
import {
  buildSuggestionProgressSummaryText,
  createSuggestionProgressSummaryModel,
} from "./SuggestionProgressSummary";
import { wireSuggestionCardInteractions } from "./suggestion-card/SuggestionCardActions";
import { createSuggestionCard } from "./suggestion-card/SuggestionCardElements";
import {
  getRequiredElement,
  setDisableTrackChangesCtaVisible,
} from "./TaskpaneUi";

export type { ResultsPanelDeps } from "./SuggestionCardRenderer.types";

/** Builds the summary sentence displayed above the rendered suggestion list. */
export function buildResultsSummary(
  suggestions: Suggestion[],
  result: ApplySuggestionsResult,
  chunkErrors: string[],
  isSelection: boolean,
): string {
  return buildSuggestionProgressSummaryText(
    createSuggestionProgressSummaryModel(suggestions, result, chunkErrors),
    isSelection,
  );
}

/** Builds a natural status-bar message for mixed apply outcomes. */
export function buildApplyStatusMessage(
  result: ApplySuggestionsResult,
  isSelection: boolean,
): string {
  const scopeSuffix = isSelection ? " (selección)" : "";
  const notFoundCount = result.failedSuggestions.filter(
    (failure) => failure.reason === "not-found",
  ).length;
  const failedOtherCount = result.failedSuggestions.length - notFoundCount;

  if (result.successCount > 0 && result.failedSuggestions.length === 0) {
    return `${result.successCount} sugerencia(s) insertada(s) como Track Changes${scopeSuffix}.`;
  }

  if (result.successCount > 0) {
    const fragments = [`${result.successCount} aplicada(s)`];
    if (notFoundCount > 0) {
      fragments.push(`${notFoundCount} no encontrada(s)`);
    }
    if (failedOtherCount > 0) {
      fragments.push(`${failedOtherCount} fallida(s)`);
    }
    return `${fragments.join(", ")}${scopeSuffix}.`;
  }

  return "Ninguna sugerencia pudo aplicarse al documento actual.";
}

/**
 * Renders the results panel showing each suggestion and its outcome.
 * Wires all card interactions (navigation, accept, reject, feedback).
 */
export function renderResultsPanel(
  suggestions: Suggestion[],
  result: ApplySuggestionsResult,
  chunkErrors: string[],
  isSelection: boolean,
  deps: ResultsPanelDeps,
): void {
  const panel = getRequiredElement("results-panel");
  const summary = getRequiredElement("results-summary");
  const list = getRequiredElement("results-list");
  const summaryModel = createSuggestionProgressSummaryModel(
    suggestions,
    result,
    chunkErrors,
  );

  summary.textContent = buildSuggestionProgressSummaryText(
    summaryModel,
    isSelection,
  );
  const uiContext = {
    summaryModel,
    summaryElement: summary,
    isSelection,
  };

  list.innerHTML = "";
  const cards = suggestions.map((suggestion) =>
    createSuggestionCard(suggestion, result.failedSuggestions),
  );

  for (const card of cards.filter((entry) => !entry.isNotFoundFailure)) {
    list.appendChild(card.li);
    if (!card.isFailed) {
      wireSuggestionCardInteractions(card.li, card.suggestion, deps, uiContext);
    }
  }

  for (const card of cards.filter((entry) => entry.isNotFoundFailure)) {
    list.appendChild(card.li);
  }

  panel.style.display = "block";
  setDisableTrackChangesCtaVisible(
    result.documentState === "ready-to-disable-track-changes",
  );
}
