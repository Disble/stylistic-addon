import type { SuggestionApplicationFailure } from "../../../domain/DocumentApplication.types";
import type { SuggestionSeverity } from "../../../domain/suggestion/Suggestion.types";
import type { ResultsPanelCardState } from "../../ResultsPanelStore";

export type SeverityBadgeColor = "danger" | "warning" | "informative" | "subtle";

const SEVERITY_LABEL: Record<SuggestionSeverity, string> = {
  high: "alta",
  medium: "media",
  low: "baja",
};

const SEVERITY_COLOR: Record<SuggestionSeverity, SeverityBadgeColor> = {
  high: "danger",
  medium: "warning",
  low: "informative",
};

/** Maps a suggestion severity to a Fluent Badge color token. */
export function getSeverityBadgeColor(severity: SuggestionSeverity): SeverityBadgeColor {
  return SEVERITY_COLOR[severity] ?? "subtle";
}

/** Returns a humanized severity label for the badge text. */
export function getSeverityLabel(severity: SuggestionSeverity): string {
  return SEVERITY_LABEL[severity] ?? severity;
}

/** Returns the failure copy for a card that never applied successfully. */
export function getFailedSuggestionCopy(failure: SuggestionApplicationFailure): string {
  return failure.reason === "not-found"
    ? `No encontrado: "${failure.suggestion.anchor}"`
    : `No se pudo aplicar: "${failure.suggestion.anchor}"`;
}

/** Whether the card represents a comment-only suggestion (no diff). */
export function isCommentOnlyCard(card: ResultsPanelCardState): boolean {
  return card.suggestion.type === "comment-only";
}
