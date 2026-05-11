import type { SuggestionApplicationFailure } from "../../../domain/DocumentApplication.types";
import type { SuggestionSeverity } from "../../../domain/suggestion/Suggestion.types";
import type { ResultsPanelCardState } from "../../ResultsPanelStore.types";
import {
  RESULT_SUGGESTION_CARD_DIACRITIC_PATTERN,
  RESULT_SUGGESTION_CARD_SEVERITY_LABEL,
} from "./ResultSuggestionCard.constants";
import type {
  CardVisualState,
  CategoryAccent,
  ResultSuggestionCardStyles,
} from "./ResultSuggestionCard.types";

/** Returns a humanized severity label for inline dot rendering. */
export function getSeverityLabel(severity: SuggestionSeverity): string {
  return RESULT_SUGGESTION_CARD_SEVERITY_LABEL[severity] ?? severity;
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
  const slug = category
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(RESULT_SUGGESTION_CARD_DIACRITIC_PATTERN, "");
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

/** Whether the card should render the original-vs-suggested diff block. */
export function shouldShowSuggestionDiff(card: ResultsPanelCardState): boolean {
  return card.suggestion.type !== "comment-only";
}

/** Whether the card should render the severity badge. */
export function shouldShowSeverity(card: ResultsPanelCardState): boolean {
  return card.isFailed !== true;
}

/** Resolves the outer card state class for one suggestion state. */
export function resolveCardStateClass(
  styles: ResultSuggestionCardStyles,
  state: CardVisualState
): string {
  if (state === "accepted") return styles.cardStateAccepted;
  if (state === "rejected") return styles.cardStateRejected;
  if (state === "failed") return styles.cardStateFailed;
  if (state === "not-found") return styles.cardStateNotFound;
  return styles.cardStatePending;
}

/** Resolves the colored stripe class for one suggestion state. */
export function resolveStripeClass(
  styles: ResultSuggestionCardStyles,
  state: CardVisualState
): string {
  if (state === "accepted") return styles.stripeAccepted;
  if (state === "rejected") return styles.stripeRejected;
  if (state === "failed") return styles.stripeFailed;
  if (state === "not-found") return styles.stripeNotFound;
  return styles.stripePending;
}

/** Resolves the category pill class for one accent slug. */
export function resolveCategoryClass(
  styles: ResultSuggestionCardStyles,
  accent: CategoryAccent
): string {
  if (accent === "grammar") return styles.categoryGrammar;
  if (accent === "spelling") return styles.categorySpelling;
  if (accent === "punctuation") return styles.categoryPunctuation;
  if (accent === "style") return styles.categoryStyle;
  return styles.categoryNeutral;
}

/** Resolves the severity dot class for one severity level. */
export function resolveSeverityDotClass(
  styles: ResultSuggestionCardStyles,
  severity: SuggestionSeverity
): string {
  if (severity === "high") return styles.severityDotHigh;
  if (severity === "medium") return styles.severityDotMedium;
  return styles.severityDotLow;
}

/** Resolves the status icon color class for one card state. */
export function resolveStatusIconClass(
  styles: ResultSuggestionCardStyles,
  state: CardVisualState
): string {
  if (state === "accepted") return styles.statusIconAccepted;
  if (state === "rejected") return styles.statusIconRejected;
  if (state === "failed") return styles.statusIconFailed;
  if (state === "not-found") return styles.statusIconNotFound;
  return "";
}
