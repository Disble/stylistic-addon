/**
 * SuggestionProgressSummary — taskpane-only progress accounting for the live
 * results summary shown above the suggestion list.
 *
 * This module intentionally stays in the presentation layer: it translates the
 * already-known apply outcome plus later user resolution outcomes into a
 * user-facing summary sentence.
 *
 * @module SuggestionProgressSummary
 */

import type { ApplySuggestionsResult } from "../domain/DocumentApplication.types";
import type { Suggestion } from "../domain/suggestion/Suggestion.types";
import type { SuggestionActionResult } from "../domain/suggestion/SuggestionResolutionWorkflow.types";

/** Terminal progress buckets for suggestions already applied to the document. */
type AppliedSuggestionProgressState =
  | "pending"
  | "resolved"
  | "needs-attention";

/** Mutable progress snapshot owned by the results panel during one render cycle. */
export interface SuggestionProgressSummaryModel {
  /** Total suggestions emitted by the analysis pipeline. */
  total: number;

  /** Suggestions successfully materialized in Word during the initial batch. */
  applied: number;

  /** Suggestions never found in the document during the initial batch. */
  notFound: number;

  /** Suggestions that failed to apply for reasons other than not-found. */
  failedOther: number;

  /** Chunk-level analysis errors preserved from the pipeline. */
  chunkErrors: number;

  /** Per-suggestion live state for suggestions that were initially applied. */
  appliedStates: Map<string, AppliedSuggestionProgressState>;
}

/**
 * Creates the initial live summary model from the pipeline completion payload.
 */
export function createSuggestionProgressSummaryModel(
  suggestions: Suggestion[],
  result: ApplySuggestionsResult,
  chunkErrors: string[],
): SuggestionProgressSummaryModel {
  const failedSuggestionIds = new Set(
    result.failedSuggestions.map((failure) => failure.suggestion.id),
  );
  const appliedStates = new Map(
    suggestions
      .filter((suggestion) => !failedSuggestionIds.has(suggestion.id))
      .map((suggestion) => [
        suggestion.id,
        "pending" as AppliedSuggestionProgressState,
      ]),
  );

  return {
    total: suggestions.length,
    applied: result.successCount,
    notFound: result.failedSuggestions.filter(
      (failure) => failure.reason === "not-found",
    ).length,
    failedOther: result.failedSuggestions.filter(
      (failure) => failure.reason !== "not-found",
    ).length,
    chunkErrors: chunkErrors.length,
    appliedStates,
  };
}

/**
 * Applies a resolution outcome to the live summary model.
 *
 * Retryable outcomes keep the suggestion in `pending`. Terminal warning states
 * such as `identity-lost` and `cc-not-found` move it to `needs-attention` so
 * the user can distinguish true progress from blocked work.
 */
export function applySuggestionProgressOutcome(
  model: SuggestionProgressSummaryModel,
  suggestionId: string,
  status: SuggestionActionResult["status"],
): void {
  if (!model.appliedStates.has(suggestionId)) {
    return;
  }

  const nextState = mapActionStatusToProgressState(status);
  if (!nextState) {
    return;
  }

  model.appliedStates.set(suggestionId, nextState);
}

/** Builds the user-facing summary sentence from the current live model. */
export function buildSuggestionProgressSummaryText(
  model: SuggestionProgressSummaryModel,
  isSelection: boolean,
): string {
  const scopePrefix = isSelection ? "Sobre selección — " : "";
  const resolved = countAppliedState(model, "resolved");
  const remaining = countAppliedState(model, "pending");
  const needsAttention = countAppliedState(model, "needs-attention");

  let summaryText = buildHumanPlanningSummary(
    scopePrefix,
    model.applied,
    resolved,
    remaining,
  );

  if (model.notFound > 0) {
    summaryText += ` ${model.notFound} no encontrada(s) en el texto.`;
  }
  if (model.failedOther > 0) {
    summaryText += ` ${model.failedOther} no pudo/pudieron aplicarse.`;
  }
  if (needsAttention > 0) {
    summaryText += ` ${needsAttention} requiere(n) revisión manual.`;
  }
  if (model.chunkErrors > 0) {
    summaryText += ` ${model.chunkErrors} fragmento(s) con error.`;
  }

  return summaryText;
}

/** Builds the primary copy focused on how much review work remains. */
function buildHumanPlanningSummary(
  scopePrefix: string,
  applied: number,
  resolved: number,
  remaining: number,
): string {
  if (applied === 0) {
    return `${scopePrefix}No hay sugerencias aplicadas para revisar.`;
  }

  if (remaining === 0) {
    if (resolved === 0) {
      return `${scopePrefix}Ya no te quedan sugerencias aplicadas por revisar.`;
    }

    return `${scopePrefix}Ya no te quedan sugerencias aplicadas por revisar. ${resolved} ${pluralize(
      resolved,
      "ya resuelta.",
      "ya resueltas.",
    )}`;
  }

  const suggestionLabel = pluralize(
    applied,
    "sugerencia aplicada",
    "sugerencias aplicadas",
  );
  const resolvedText =
    resolved === 0
      ? "Todavía no resolviste ninguna."
      : `${resolved} ${pluralize(resolved, "ya resuelta.", "ya resueltas.")}`;

  return `${scopePrefix}Te faltan ${remaining} de ${applied} ${suggestionLabel} por revisar. ${resolvedText}`;
}

/** Counts how many applied suggestions currently belong to one progress bucket. */
function countAppliedState(
  model: SuggestionProgressSummaryModel,
  state: AppliedSuggestionProgressState,
): number {
  return Array.from(model.appliedStates.values()).filter(
    (currentState) => currentState === state,
  ).length;
}

/** Returns the singular or plural copy for a count. */
function pluralize(count: number, singular: string, plural: string): string {
  return count === 1 ? singular : plural;
}

/** Maps a document resolution outcome to a live summary bucket. */
function mapActionStatusToProgressState(
  status: SuggestionActionResult["status"],
): AppliedSuggestionProgressState | null {
  switch (status) {
    case "accepted":
    case "rejected":
      return "resolved";
    case "identity-lost":
    case "ambiguous-location":
    case "mixed-group":
    case "cc-not-found":
      return "needs-attention";
    case "unobservable":
    case "not-found":
    case "error":
      return "pending";
  }
}
