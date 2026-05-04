/* global document */

import type { SuggestionState } from "../../domain/suggestion/Suggestion.types";
import type { SuggestionResolutionMediatorResult } from "../../domain/suggestion/SuggestionResolutionWorkflow.types";
import type { SuggestionResolutionUiContext } from "../SuggestionCardRenderer.types";
import {
  applySuggestionProgressOutcome,
  buildSuggestionProgressSummaryText,
} from "../SuggestionProgressSummary";
import { appendNote, setDisableTrackChangesCtaVisible, showStatus } from "../TaskpaneUi";
import { moveSuggestionCardToEnd } from "./SuggestionCardList";

/** Updates the DOM for a suggestion card based on the SM's terminal state. */
export function applySuggestionCardState(
  li: HTMLElement,
  state: SuggestionState,
  acceptBtn: HTMLButtonElement | null,
  rejectBtn: HTMLButtonElement | null,
  errorMessage?: string
): void {
  switch (state) {
    case "accepted":
    case "rejected":
      li.querySelector(".result-actions")?.remove();
      li.classList.add(`result-${state}`);
      moveSuggestionCardToEnd(li);
      break;

    case "identity-lost":
    case "ambiguous-location":
    case "mixed-group":
      li.querySelector(".result-actions")?.remove();
      li.classList.add(`result-${state}`);
      appendNote(
        li,
        state === "identity-lost"
          ? "(metadata inconsistente; reanalizá la sugerencia)"
          : "(resolución ambigua; reanalizá la sugerencia)",
        `result-${state}-note`
      );
      showStatus(
        errorMessage ??
          "La identidad persistida de la sugerencia no permite resolver con seguridad.",
        "error"
      );
      moveSuggestionCardToEnd(li);
      break;

    case "unobservable":
      if (acceptBtn) acceptBtn.disabled = false;
      if (rejectBtn) rejectBtn.disabled = false;
      showStatus(
        errorMessage ?? "No se pudo confirmar el estado de la sugerencia en Word. Reintentá.",
        "error"
      );
      break;

    case "error":
      if (acceptBtn) acceptBtn.disabled = false;
      if (rejectBtn) rejectBtn.disabled = false;
      showStatus(errorMessage ?? "Error desconocido al resolver sugerencia", "error");
      break;

    default:
      break;
  }
}

/** Applies shared taskpane consequences after a workflow-owned resolution. */
export function applyResolutionWorkflowUi(result: SuggestionResolutionMediatorResult): void {
  setDisableTrackChangesCtaVisible(result.taskpaneState.showDisableTrackChangesCta);
  const cleanupSection = document.getElementById("cleanup-section");
  if (cleanupSection) {
    cleanupSection.style.display = result.taskpaneState.showCleanupSection ? "block" : "none";
  }
}

/** Updates the live summary text after one suggestion resolution outcome. */
export function updateResultsSummaryAfterResolution(
  uiContext: SuggestionResolutionUiContext,
  suggestionId: string,
  result: SuggestionResolutionMediatorResult
): void {
  applySuggestionProgressOutcome(uiContext.summaryModel, suggestionId, result.status);
  uiContext.summaryElement.textContent = buildSuggestionProgressSummaryText(
    uiContext.summaryModel,
    uiContext.isSelection
  );
}
