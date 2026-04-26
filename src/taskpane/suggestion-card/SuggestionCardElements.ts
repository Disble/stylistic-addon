/* global document */

import type { SuggestionApplicationFailure } from "../../domain/DocumentApplication.types";
import type { Suggestion } from "../../domain/suggestion/Suggestion.types";
import type { RenderedSuggestionCard } from "../SuggestionCardRenderer.types";

/** Creates the metadata row for one suggestion card. */
export function createSuggestionMetaRow(
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
export function appendFailedCardContent(
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
export function createActionButton(
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
export function appendActionableCardContent(
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
export function createSuggestionCard(
  suggestion: Suggestion,
  failedSuggestions: SuggestionApplicationFailure[],
): RenderedSuggestionCard {
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
