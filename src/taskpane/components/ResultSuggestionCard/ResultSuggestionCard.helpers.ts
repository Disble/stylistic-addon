import type { SuggestionApplicationFailure } from "../../../domain/DocumentApplication.types";
import type { SuggestionSeverity } from "../../../domain/suggestion/Suggestion.types";
import type { ResultsPanelCardState } from "../../ResultsPanelStore";

export type CardVisualState = "pending" | "accepted" | "rejected" | "failed" | "not-found";

export type CategoryAccent = "grammar" | "spelling" | "punctuation" | "style" | "neutral";

const SEVERITY_LABEL: Record<SuggestionSeverity, string> = {
  high: "alta",
  medium: "media",
  low: "baja",
};

const DIACRITIC_PATTERN = /[̀-ͯ]/g;

/** Returns a humanized severity label for inline dot rendering. */
export function getSeverityLabel(severity: SuggestionSeverity): string {
  return SEVERITY_LABEL[severity] ?? severity;
}

/** Resolves the visual state used to color the stripe and background tint. */
export function resolveCardVisualState(card: ResultsPanelCardState): CardVisualState {
  if (card.cardGroup === "not-found") return "not-found";
  if (card.isFailed) return "failed";
  if (card.state === "accepted") return "accepted";
  if (card.state === "rejected") return "rejected";
  return "pending";
}

/** Maps the raw suggestion category to a stable accent slug (tolerant of accents). */
export function getCategoryAccent(category: string): CategoryAccent {
  const slug = category.trim().toLowerCase().normalize("NFD").replace(DIACRITIC_PATTERN, "");
  if (slug.startsWith("gramatic") || slug.startsWith("grammar")) return "grammar";
  if (slug.startsWith("ortograf") || slug.startsWith("spelling")) return "spelling";
  if (slug.startsWith("puntuac") || slug.startsWith("punctuation")) return "punctuation";
  if (slug.startsWith("estilo") || slug.startsWith("style")) return "style";
  return "neutral";
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
