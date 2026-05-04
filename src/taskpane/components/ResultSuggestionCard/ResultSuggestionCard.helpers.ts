import type { SuggestionApplicationFailure } from "../../../domain/DocumentApplication.types";
import type { ResultsPanelCardState } from "../../ResultsPanelStore";

/** Returns the rendered card CSS classes from the current view state. */
export function getResultSuggestionCardClassName(card: ResultsPanelCardState): string {
  const classes = ["suggestion-card"];

  if (card.state !== "pending" && card.state !== "resolving") {
    classes.push(`result-${card.state}`);
  }

  return classes.join(" ");
}

/** Returns the failure copy for a card that never applied successfully. */
export function getFailedSuggestionCopy(failure: SuggestionApplicationFailure): string {
  return failure.reason === "not-found"
    ? `No encontrado: "${failure.suggestion.anchor}"`
    : `No se pudo aplicar: "${failure.suggestion.anchor}"`;
}

/** Whether the card should render the comment-only badge. */
export function isCommentOnlyCard(card: ResultsPanelCardState): boolean {
  return card.suggestion.type === "comment-only";
}
