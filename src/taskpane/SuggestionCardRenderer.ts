/**
 * SuggestionCardRenderer — taskpane-facing facade for results-panel publishing.
 *
 * React now owns the concrete results rendering. This module keeps only the
 * public presentation helpers used by the taskpane composition root and publishes
 * pipeline-completion payloads into `ResultsPanelStore`.
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
import { setResultsPanelData } from "./ResultsPanelStore";

/** Builds the summary sentence displayed above the rendered suggestion list. */
export function buildResultsSummary(
  suggestions: Suggestion[],
  result: ApplySuggestionsResult,
  chunkErrors: string[],
  isSelection: boolean
): string {
  return buildSuggestionProgressSummaryText(
    createSuggestionProgressSummaryModel(suggestions, result, chunkErrors),
    isSelection
  );
}

/** Builds a natural status-bar message for mixed apply outcomes. */
export function buildApplyStatusMessage(
  result: ApplySuggestionsResult,
  isSelection: boolean
): string {
  const scopeSuffix = isSelection ? " (selección)" : "";
  const notFoundCount = result.failedSuggestions.filter(
    (failure) => failure.reason === "not-found"
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
  deps: ResultsPanelDeps
): void {
  setResultsPanelData(suggestions, result, chunkErrors, isSelection, deps);
}
