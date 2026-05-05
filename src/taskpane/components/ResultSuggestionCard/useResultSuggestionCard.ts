import { makeStyles, tokens } from "@fluentui/react-components";
import {
  getSeverityBadgeColor,
  getSeverityLabel,
  isCommentOnlyCard,
} from "./ResultSuggestionCard.helpers";
import type {
  ResultSuggestionCardClasses,
  ResultSuggestionCardViewModel,
} from "./ResultSuggestionCard.types";
import type { ResultsPanelCardState } from "../../ResultsPanelStore";

const SEVERITY_STRIPE_COLOR: Record<string, string> = {
  high: tokens.colorStatusDangerBorder2,
  medium: tokens.colorStatusWarningBorder2,
  low: tokens.colorBrandStroke1,
};

const useResultSuggestionCardStyles = makeStyles({
  root: {
    listStyle: "none",
    margin: 0,
    marginBottom: tokens.spacingVerticalM,
    padding: 0,
  },
  card: {
    position: "relative",
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
    rowGap: tokens.spacingVerticalS,
    overflow: "hidden",
  },
  severityStripeBase: {
    position: "absolute",
    top: 0,
    bottom: 0,
    left: 0,
    width: "4px",
  },
  severityStripeHigh: {
    backgroundColor: SEVERITY_STRIPE_COLOR.high,
  },
  severityStripeMedium: {
    backgroundColor: SEVERITY_STRIPE_COLOR.medium,
  },
  severityStripeLow: {
    backgroundColor: SEVERITY_STRIPE_COLOR.low,
  },
  badgeRow: {
    display: "flex",
    flexWrap: "wrap",
    columnGap: tokens.spacingHorizontalXS,
    rowGap: tokens.spacingVerticalXS,
    alignItems: "center",
  },
  diff: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalXXS,
    paddingTop: tokens.spacingVerticalXS,
  },
  diffArrow: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  original: {
    color: tokens.colorPaletteRedForeground1,
    textDecorationLine: "line-through",
    overflowWrap: "anywhere",
  },
  suggested: {
    color: tokens.colorPaletteGreenForeground1,
    overflowWrap: "anywhere",
  },
  justification: {
    color: tokens.colorNeutralForeground2,
    overflowWrap: "anywhere",
    margin: 0,
  },
  clickableArea: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalXS,
    background: "transparent",
    border: "none",
    padding: 0,
    margin: 0,
    textAlign: "left",
    cursor: "pointer",
    color: "inherit",
    width: "100%",
  },
  failureLabel: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: tokens.fontWeightSemibold,
  },
  failureDetail: {
    color: tokens.colorNeutralForeground3,
  },
  note: {
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
  },
  footer: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalS,
    paddingTop: tokens.spacingVerticalXS,
  },
  actions: {
    display: "flex",
    columnGap: tokens.spacingHorizontalXS,
    alignItems: "center",
  },
  primaryAction: {
    flexGrow: 1,
  },
  feedbackTextarea: {
    width: "100%",
  },
});

/** Resolves the Fluent classes and derived view-model for one suggestion card. */
export function useResultSuggestionCard(
  card: ResultsPanelCardState
): ResultSuggestionCardViewModel {
  const styles = useResultSuggestionCardStyles();
  const isCommentOnly = isCommentOnlyCard(card);
  const severityStripeClass = resolveSeverityStripeClass(styles, card.suggestion.severity);

  const classes: ResultSuggestionCardClasses = {
    root: styles.root,
    card: styles.card,
    severityStripe: `${styles.severityStripeBase} ${severityStripeClass}`.trim(),
    badgeRow: styles.badgeRow,
    diff: styles.diff,
    diffArrow: styles.diffArrow,
    original: styles.original,
    suggested: styles.suggested,
    justification: styles.justification,
    clickableArea: styles.clickableArea,
    failureLabel: styles.failureLabel,
    failureDetail: styles.failureDetail,
    note: styles.note,
    footer: styles.footer,
    actions: styles.actions,
    primaryAction: styles.primaryAction,
    feedbackTextarea: styles.feedbackTextarea,
  };

  return {
    classes,
    isCommentOnly,
    severityColor: getSeverityBadgeColor(card.suggestion.severity),
    severityLabel: getSeverityLabel(card.suggestion.severity),
    acceptLabel: isCommentOnly ? "Entendido" : "Aceptar",
    rejectLabel: isCommentOnly ? "Ignorar" : "Rechazar",
    acceptAriaLabel: isCommentOnly ? "Marcar comentario como entendido" : "Aceptar sugerencia",
    rejectAriaLabel: isCommentOnly ? "Ignorar comentario" : "Rechazar sugerencia",
    feedbackToggleAriaLabel: card.feedbackOpen ? "Cerrar feedback" : "Dejar feedback",
  };
}

function resolveSeverityStripeClass(
  styles: ReturnType<typeof useResultSuggestionCardStyles>,
  severity: ResultsPanelCardState["suggestion"]["severity"]
): string {
  if (severity === "high") return styles.severityStripeHigh;
  if (severity === "medium") return styles.severityStripeMedium;
  if (severity === "low") return styles.severityStripeLow;
  return "";
}
