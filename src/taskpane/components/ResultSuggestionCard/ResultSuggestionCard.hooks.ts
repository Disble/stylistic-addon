import {
  getCategoryAccent,
  getSeverityLabel,
  isCommentOnlyCard,
  resolveCardStateClass,
  resolveCategoryClass,
  resolveCardVisualState,
  resolveSeverityDotClass,
  resolveStatusIconClass,
  resolveStripeClass,
} from "./ResultSuggestionCard.helpers";
import { useResultSuggestionCardStyles } from "./ResultSuggestionCard.styles";
import type {
  ResultSuggestionCardClasses,
  ResultSuggestionCardViewModel,
} from "./ResultSuggestionCard.types";
import type { ResultsPanelCardState } from "../../ResultsPanelStore.types";

/** Resolves the Fluent classes and derived view-model for one suggestion card. */
export function useResultSuggestionCard(
  card: ResultsPanelCardState
): ResultSuggestionCardViewModel {
  const styles = useResultSuggestionCardStyles();
  const isCommentOnly = isCommentOnlyCard(card);
  const cardVisualState = resolveCardVisualState(card);
  const categoryAccent = getCategoryAccent(card.suggestion.category);

  const classes: ResultSuggestionCardClasses = {
    root: styles.root,
    card: `${styles.card} ${resolveCardStateClass(styles, cardVisualState)}`.trim(),
    stateStripe: `${styles.stripeBase} ${resolveStripeClass(styles, cardVisualState)}`.trim(),
    badgeRow: styles.badgeRow,
    categoryPill:
      `${styles.categoryPillBase} ${resolveCategoryClass(styles, categoryAccent)}`.trim(),
    severityIndicator: styles.severityIndicator,
    severityDot:
      `${styles.severityDotBase} ${resolveSeverityDotClass(styles, card.suggestion.severity)}`.trim(),
    severityLabel: styles.severityLabel,
    statusIcon:
      `${styles.statusIconBase} ${resolveStatusIconClass(styles, cardVisualState)}`.trim(),
    diff: styles.diff,
    original: styles.original,
    suggested: styles.suggested,
    justification: styles.justification,
    clickableArea: styles.clickableArea,
    failureLabel: styles.failureLabel,
    failureDetail: styles.failureDetail,
    note: styles.note,
    footer: styles.footer,
    actions: styles.actions,
    feedbackTextarea: styles.feedbackTextarea,
  };

  return {
    classes,
    isCommentOnly,
    cardVisualState,
    categoryAccent,
    severityLabel: getSeverityLabel(card.suggestion.severity),
    acceptLabel: isCommentOnly ? "Entendido" : "Aceptar",
    rejectLabel: isCommentOnly ? "Ignorar" : "Rechazar",
    acceptAriaLabel: isCommentOnly ? "Marcar comentario como entendido" : "Aceptar sugerencia",
    rejectAriaLabel: isCommentOnly ? "Ignorar comentario" : "Rechazar sugerencia",
    feedbackToggleAriaLabel: card.feedbackOpen ? "Cerrar feedback" : "Dejar feedback",
  };
}
