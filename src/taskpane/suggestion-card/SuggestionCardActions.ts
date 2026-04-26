import type { Suggestion } from "../../domain/suggestion/Suggestion.types";
import {
  mapResultStatusToState,
  SuggestionStateMachine,
} from "../../domain/suggestion/SuggestionStateMachine";
import type {
  ResultsPanelDeps,
  SuggestionActionButtons,
  SuggestionResolutionUiContext,
} from "../SuggestionCardRenderer.types";
import { appendNote } from "../TaskpaneUi";
import {
  getSuggestionFeedbackComment,
  wireSuggestionFeedbackToggle,
} from "./SuggestionCardFeedback";
import { moveSuggestionCardToEnd } from "./SuggestionCardList";
import {
  applyResolutionWorkflowUi,
  applySuggestionCardState,
  updateResultsSummaryAfterResolution,
} from "./SuggestionCardStateRenderer";

/** Handles an accept action for one suggestion card. */
export async function handleAcceptSuggestion(
  suggestion: Suggestion,
  li: HTMLElement,
  buttons: SuggestionActionButtons,
  sm: SuggestionStateMachine,
  deps: ResultsPanelDeps,
  uiContext: SuggestionResolutionUiContext,
): Promise<void> {
  if (!sm.canTransition("resolving")) return;

  sm.transition("resolving");
  if (buttons.acceptBtn) buttons.acceptBtn.disabled = true;
  if (buttons.rejectBtn) buttons.rejectBtn.disabled = true;

  const result = await deps.acceptSuggestion(
    suggestion,
    getSuggestionFeedbackComment(li),
  );

  if (result.status === "cc-not-found") {
    sm.transition("error");
    li.querySelector(".result-actions")?.remove();
    li.classList.add("result-cc-not-found");
    appendNote(li, "(aplicación falló)", "result-cc-not-found-note");
    moveSuggestionCardToEnd(li);
    return;
  }

  const targetState = mapResultStatusToState(result.status);
  sm.transition(targetState);
  applySuggestionCardState(
    li,
    sm.state,
    buttons.acceptBtn,
    buttons.rejectBtn,
    result.error,
  );
  applyResolutionWorkflowUi(result);
  updateResultsSummaryAfterResolution(uiContext, suggestion.id, result);
}

/** Handles a reject action for one suggestion card. */
export async function handleRejectSuggestion(
  suggestion: Suggestion,
  li: HTMLElement,
  buttons: SuggestionActionButtons,
  sm: SuggestionStateMachine,
  deps: ResultsPanelDeps,
  uiContext: SuggestionResolutionUiContext,
): Promise<void> {
  if (!sm.canTransition("resolving")) return;

  sm.transition("resolving");
  if (buttons.acceptBtn) buttons.acceptBtn.disabled = true;
  if (buttons.rejectBtn) buttons.rejectBtn.disabled = true;

  const result = await deps.rejectSuggestion(
    suggestion,
    getSuggestionFeedbackComment(li),
  );

  if (result.status === "cc-not-found") {
    sm.transition("error");
    li.querySelector(".result-actions")?.remove();
    li.classList.add("result-cc-not-found");
    appendNote(li, "(aplicación falló)", "result-cc-not-found-note");
    moveSuggestionCardToEnd(li);
    return;
  }

  const targetState = mapResultStatusToState(result.status);
  sm.transition(targetState);
  applySuggestionCardState(
    li,
    sm.state,
    buttons.acceptBtn,
    buttons.rejectBtn,
    result.error,
  );
  applyResolutionWorkflowUi(result);
  updateResultsSummaryAfterResolution(uiContext, suggestion.id, result);
}

/** Wires per-card interaction handlers for navigation, feedback, accept and reject. */
export function wireSuggestionCardInteractions(
  li: HTMLLIElement,
  suggestion: Suggestion,
  deps: ResultsPanelDeps,
  uiContext: SuggestionResolutionUiContext,
): void {
  const clickableEl = li.querySelector(
    ".card-clickable-area",
  ) as HTMLElement | null;
  if (clickableEl) {
    clickableEl.addEventListener("click", () => {
      void deps.navigateToText(suggestion);
    });
  }

  const acceptBtnEl = li.querySelector(
    '[data-action="accept"]',
  ) as HTMLButtonElement | null;
  const rejectBtnEl = li.querySelector(
    '[data-action="reject"]',
  ) as HTMLButtonElement | null;

  wireSuggestionFeedbackToggle(li);

  const sm = new SuggestionStateMachine();
  const buttons: SuggestionActionButtons = {
    acceptBtn: acceptBtnEl,
    rejectBtn: rejectBtnEl,
  };

  if (acceptBtnEl) {
    acceptBtnEl.addEventListener("click", () =>
      handleAcceptSuggestion(suggestion, li, buttons, sm, deps, uiContext),
    );
  }
  if (rejectBtnEl) {
    rejectBtnEl.addEventListener("click", () =>
      handleRejectSuggestion(suggestion, li, buttons, sm, deps, uiContext),
    );
  }
}
