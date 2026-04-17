/* global document */

/**
 * SuggestionCardRenderer — creates, renders, and wires suggestion cards.
 *
 * Extracted from `taskpane.ts` to keep the Composition Root focused on
 * initialization, pipeline orchestration, and top-level event handlers.
 *
 * All business-layer dependencies (document port, mediator) are injected
 * via `ResultsPanelDeps` to avoid module-level coupling.
 *
 * @module SuggestionCardRenderer
 */

import {
  mapResultStatusToState,
  SuggestionStateMachine,
} from "../domain/suggestion/SuggestionStateMachine";
import type {
  ApplySuggestionsResult,
  Suggestion,
  SuggestionApplicationFailure,
  SuggestionResolutionMediatorResult,
  SuggestionState,
} from "../domain/types";
import { SUGGESTION_CARD_REORDER_ANIMATION_MS } from "../infrastructure/config";
import {
  applySuggestionProgressOutcome,
  buildSuggestionProgressSummaryText,
  createSuggestionProgressSummaryModel,
  type SuggestionProgressSummaryModel,
} from "./SuggestionProgressSummary";
import {
  appendNote,
  getRequiredElement,
  setDisableTrackChangesCtaVisible,
  showStatus,
} from "./TaskpaneUi";

// ---------------------------------------------------------------------------
// Dependencies — injected by the caller, never imported at module level
// ---------------------------------------------------------------------------

/** Business-layer capabilities needed to render and interact with cards. */
export type ResultsPanelDeps = {
  navigateToText: (target: Suggestion | string) => Promise<void>;
  acceptSuggestion: (
    suggestion: Suggestion,
    comment?: string,
  ) => Promise<SuggestionResolutionMediatorResult>;
  rejectSuggestion: (
    suggestion: Suggestion,
    comment?: string,
  ) => Promise<SuggestionResolutionMediatorResult>;
};

/** Shared UI state required while resolving one rendered suggestion card. */
type SuggestionResolutionUiContext = {
  summaryModel: SuggestionProgressSummaryModel;
  summaryElement: HTMLElement;
  isSelection: boolean;
};

/** The pair of action buttons owned by one suggestion card. */
type SuggestionActionButtons = {
  acceptBtn: HTMLButtonElement | null;
  rejectBtn: HTMLButtonElement | null;
};

// ---------------------------------------------------------------------------
// Pure summary builders
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Card element builders
// ---------------------------------------------------------------------------

/** Creates the metadata row for one suggestion card. */
function createSuggestionMetaRow(
  suggestion: Suggestion,
  isFailed: boolean,
  isCommentOnly: boolean,
): HTMLDivElement {
  const meta = document.createElement("div");
  meta.className = "card-meta";

  const catBadge = document.createElement("span");
  catBadge.className = "result-category";
  catBadge.textContent = suggestion.category;
  meta.appendChild(catBadge);

  if (isFailed) {
    return meta;
  }

  const sevBadge = document.createElement("span");
  sevBadge.className = `result-severity result-severity--${suggestion.severity}`;
  sevBadge.textContent = suggestion.severity;
  meta.appendChild(sevBadge);

  if (isCommentOnly) {
    const typeBadge = document.createElement("span");
    typeBadge.className = "result-type-badge result-type-badge--comment";
    typeBadge.textContent = "comentario";
    meta.appendChild(typeBadge);
  }

  return meta;
}

/** Renders the failed-state content for one suggestion card. */
function appendFailedCardContent(
  li: HTMLLIElement,
  failure: SuggestionApplicationFailure,
): void {
  const failedSpan = document.createElement("span");
  failedSpan.className = "result-failed";
  failedSpan.textContent =
    failure.reason === "not-found"
      ? `No encontrado: "${failure.suggestion.anchor}"`
      : `No se pudo aplicar: "${failure.suggestion.anchor}"`;
  li.appendChild(failedSpan);

  const justSpan = document.createElement("span");
  justSpan.className = "result-justification";
  justSpan.textContent = failure.suggestion.justification;
  li.appendChild(justSpan);

  if (failure.reason !== "not-found") {
    const detailSpan = document.createElement("span");
    detailSpan.className = "result-failure-detail";
    detailSpan.textContent = failure.message;
    li.appendChild(detailSpan);
  }
}

/** Creates one action button for accept/reject/feedback actions. */
function createActionButton(
  action: "accept" | "reject" | "feedback",
  suggestionId: string,
  isCommentOnly: boolean,
): HTMLButtonElement {
  const button = document.createElement("button");
  button.dataset.action = action;

  if (action === "feedback") {
    button.className = "feedback-btn";
    button.setAttribute("aria-label", "Dejar feedback");
    button.textContent = "💬";
    return button;
  }

  button.className = isCommentOnly
    ? "result-action-btn result-action-btn--text"
    : "result-action-btn";
  button.dataset.suggestionId = suggestionId;

  if (action === "accept") {
    button.setAttribute("aria-label", "Aceptar sugerencia");
    button.textContent = isCommentOnly ? "Entendido" : "✓";
  } else {
    button.setAttribute("aria-label", "Rechazar sugerencia");
    button.textContent = isCommentOnly ? "Ignorar" : "✗";
  }

  return button;
}

/** Renders the actionable content (diff, justification, actions, accordion). */
function appendActionableCardContent(
  li: HTMLLIElement,
  suggestion: Suggestion,
  isCommentOnly: boolean,
): void {
  const clickable = document.createElement("div");
  clickable.className = "card-clickable-area";

  if (!isCommentOnly) {
    const diff = document.createElement("div");
    diff.className = "card-diff";

    const origSpan = document.createElement("span");
    origSpan.className = "result-original";
    origSpan.textContent = suggestion.anchor;
    diff.appendChild(origSpan);

    const arrowSpan = document.createElement("span");
    arrowSpan.className = "result-arrow";
    arrowSpan.textContent = " -> ";
    diff.appendChild(arrowSpan);

    const sugSpan = document.createElement("span");
    sugSpan.className = "result-suggested";
    sugSpan.textContent = suggestion.suggestedText ?? "";
    diff.appendChild(sugSpan);

    clickable.appendChild(diff);
  }

  const justSpan = document.createElement("span");
  justSpan.className = "result-justification";
  justSpan.textContent = suggestion.justification;
  clickable.appendChild(justSpan);
  li.appendChild(clickable);

  const footer = document.createElement("div");
  footer.className = "card-footer";

  const actionsSpan = document.createElement("span");
  actionsSpan.className = "result-actions";
  actionsSpan.appendChild(
    createActionButton("accept", suggestion.id, isCommentOnly),
  );
  actionsSpan.appendChild(
    createActionButton("reject", suggestion.id, isCommentOnly),
  );
  actionsSpan.appendChild(
    createActionButton("feedback", suggestion.id, isCommentOnly),
  );

  footer.appendChild(actionsSpan);
  li.appendChild(footer);

  const accordion = document.createElement("div");
  accordion.className = "feedback-accordion";
  const textarea = document.createElement("textarea");
  textarea.className = "feedback-textarea";
  textarea.setAttribute("placeholder", "Comentario opcional...");
  accordion.appendChild(textarea);
  li.appendChild(accordion);
}

/** Builds one suggestion card and returns whether it is in failed state. */
function createSuggestionCard(
  suggestion: Suggestion,
  failedSuggestions: SuggestionApplicationFailure[],
): {
  li: HTMLLIElement;
  isFailed: boolean;
  isNotFoundFailure: boolean;
  suggestion: Suggestion;
} {
  const failure = failedSuggestions.find(
    (f) => f.suggestion.id === suggestion.id,
  );
  const isFailed = Boolean(failure);
  const isNotFoundFailure = failure?.reason === "not-found";
  const isCommentOnly = suggestion.type === "comment-only";

  const li = document.createElement("li");
  li.className = "suggestion-card";
  li.dataset.severity = suggestion.severity;
  li.dataset.cardGroup = isNotFoundFailure ? "not-found" : "active";
  li.appendChild(createSuggestionMetaRow(suggestion, isFailed, isCommentOnly));

  if (failure) {
    appendFailedCardContent(li, failure);
    return { li, isFailed: true, isNotFoundFailure, suggestion };
  }

  appendActionableCardContent(li, suggestion, isCommentOnly);
  return { li, isFailed: false, isNotFoundFailure: false, suggestion };
}

// ---------------------------------------------------------------------------
// Card state management
// ---------------------------------------------------------------------------

/**
 * Runs a FLIP-style animation for suggestion-list reordering when the host DOM
 * provides layout APIs. Falls back to an immediate reorder in test/fake DOM.
 */
function animateSuggestionListReorder(
  parent: HTMLElement,
  reorder: () => void,
): void {
  const cardsBefore = Array.from(parent.children) as HTMLElement[];
  const canAnimate =
    typeof globalThis.requestAnimationFrame === "function" &&
    cardsBefore.every(
      (card) => typeof card.getBoundingClientRect === "function",
    );

  if (!canAnimate) {
    reorder();
    return;
  }

  const firstRects = new Map(
    cardsBefore.map((card) => [card, card.getBoundingClientRect()]),
  );

  reorder();

  const cardsAfter = Array.from(parent.children) as HTMLElement[];
  for (const card of cardsAfter) {
    const firstRect = firstRects.get(card);
    if (!firstRect) {
      continue;
    }

    const lastRect = card.getBoundingClientRect();
    const deltaY = firstRect.top - lastRect.top;
    if (deltaY === 0) {
      continue;
    }

    card.style.transition = "none";
    card.style.transform = `translateY(${deltaY}px)`;
  }

  globalThis.requestAnimationFrame(() => {
    for (const card of cardsAfter) {
      const firstRect = firstRects.get(card);
      if (!firstRect) {
        continue;
      }

      const lastRect = card.getBoundingClientRect();
      const deltaY = firstRect.top - lastRect.top;
      if (deltaY === 0) {
        continue;
      }

      card.style.transition = `transform ${SUGGESTION_CARD_REORDER_ANIMATION_MS}ms ease`;
      card.style.transform = "";
    }

    globalThis.setTimeout(() => {
      for (const card of cardsAfter) {
        card.style.transition = "";
        card.style.transform = "";
      }
    }, SUGGESTION_CARD_REORDER_ANIMATION_MS);
  });
}

/**
 * Returns the first card that represents a not-found failure.
 *
 * Those cards must remain grouped at the very bottom of the list.
 */
function getFirstNotFoundCard(parent: HTMLElement): HTMLElement | null {
  return (
    (Array.from(parent.children) as HTMLElement[]).find(
      (card) => card.dataset.cardGroup === "not-found",
    ) ?? null
  );
}

/**
 * Moves a terminally processed suggestion card to the end of the actionable
 * zone while keeping all "No encontrado" cards at the absolute bottom.
 *
 * This keeps the first visible slot reserved for the next actionable card,
 * matching the desired "stack" workflow in the taskpane.
 */
function moveSuggestionCardToEnd(li: HTMLElement): void {
  const parent = li.parentElement;
  if (!parent) {
    return;
  }

  li.dataset.cardGroup = "processed";

  animateSuggestionListReorder(parent, () => {
    const firstNotFoundCard = getFirstNotFoundCard(parent);
    if (firstNotFoundCard && firstNotFoundCard !== li) {
      firstNotFoundCard.before(li);
      return;
    }

    parent.appendChild(li);
  });
}

/**
 * Updates the DOM for a suggestion card based on the SM's terminal state.
 */
function applySuggestionCardState(
  li: HTMLElement,
  state: SuggestionState,
  acceptBtn: HTMLButtonElement | null,
  rejectBtn: HTMLButtonElement | null,
  errorMessage?: string,
): void {
  switch (state) {
    case "accepted":
    case "rejected":
      li.querySelector(".result-actions")?.remove();
      li.classList.add(`result-${state}`);
      moveSuggestionCardToEnd(li);
      break;

    case "already-resolved":
      li.querySelector(".result-actions")?.remove();
      li.classList.add("result-already-resolved");
      appendNote(li, "(ya resuelto)", "result-already-resolved-note");
      moveSuggestionCardToEnd(li);
      break;

    case "identity-lost":
      li.querySelector(".result-actions")?.remove();
      li.classList.add("result-identity-lost");
      appendNote(
        li,
        "(metadata inconsistente; reanalizá la sugerencia)",
        "result-identity-lost-note",
      );
      showStatus(
        errorMessage ??
          "La identidad persistida de la sugerencia quedó inconsistente en Word.",
        "error",
      );
      moveSuggestionCardToEnd(li);
      break;

    case "unobservable":
      if (acceptBtn) acceptBtn.disabled = false;
      if (rejectBtn) rejectBtn.disabled = false;
      showStatus(
        errorMessage ??
          "No se pudo confirmar el estado de la sugerencia en Word. Reintentá.",
        "error",
      );
      break;

    case "error":
      if (acceptBtn) acceptBtn.disabled = false;
      if (rejectBtn) rejectBtn.disabled = false;
      showStatus(
        errorMessage ?? "Error desconocido al resolver sugerencia",
        "error",
      );
      break;

    default:
      break;
  }
}

/** Returns the optional free-text feedback comment associated with a card. */
function getSuggestionFeedbackComment(li: HTMLElement): string | undefined {
  const textarea = li.querySelector(".feedback-textarea") as
    | (HTMLTextAreaElement & { value?: string })
    | null;
  const commentText = textarea?.value?.trim();
  return commentText && commentText.length > 0 ? commentText : undefined;
}

/**
 * Applies shared taskpane consequences after a workflow-owned resolution.
 */
function applyResolutionWorkflowUi(
  result: SuggestionResolutionMediatorResult,
): void {
  setDisableTrackChangesCtaVisible(
    result.taskpaneState.showDisableTrackChangesCta,
  );
  const cleanupSection = document.getElementById("cleanup-section");
  if (cleanupSection) {
    cleanupSection.style.display = result.taskpaneState.showCleanupSection
      ? "block"
      : "none";
  }
}

/** Updates the live summary text after one suggestion resolution outcome. */
function updateResultsSummaryAfterResolution(
  summaryModel: SuggestionProgressSummaryModel,
  summaryElement: HTMLElement,
  suggestionId: string,
  result: SuggestionResolutionMediatorResult,
  isSelection: boolean,
): void {
  applySuggestionProgressOutcome(summaryModel, suggestionId, result.status);
  summaryElement.textContent = buildSuggestionProgressSummaryText(
    summaryModel,
    isSelection,
  );
}

// ---------------------------------------------------------------------------
// Accept / Reject handlers
// ---------------------------------------------------------------------------

async function handleAcceptSuggestion(
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
  updateResultsSummaryAfterResolution(
    uiContext.summaryModel,
    uiContext.summaryElement,
    suggestion.id,
    result,
    uiContext.isSelection,
  );
}

async function handleRejectSuggestion(
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
  updateResultsSummaryAfterResolution(
    uiContext.summaryModel,
    uiContext.summaryElement,
    suggestion.id,
    result,
    uiContext.isSelection,
  );
}

// ---------------------------------------------------------------------------
// Card interaction wiring
// ---------------------------------------------------------------------------

/** Wires per-card interaction handlers for navigation, feedback, accept and reject. */
function wireSuggestionCardInteractions(
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
  const feedbackBtnEl = li.querySelector(
    '[data-action="feedback"]',
  ) as HTMLButtonElement | null;
  const accordionEl = li.querySelector(
    ".feedback-accordion",
  ) as HTMLElement | null;

  if (feedbackBtnEl && accordionEl) {
    feedbackBtnEl.addEventListener("click", () => {
      accordionEl.classList.toggle("feedback-accordion--open");
    });
  }

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

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/**
 * Renders the results panel showing each suggestion and its outcome.
 * Wires all card interactions (navigation, accept, reject, feedback).
 *
 * Business-layer dependencies are injected via `deps` to keep this module
 * decoupled from the Composition Root's module-level state.
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
  const uiContext: SuggestionResolutionUiContext = {
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
