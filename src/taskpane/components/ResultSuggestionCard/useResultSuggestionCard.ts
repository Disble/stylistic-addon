import { makeStyles, tokens } from "@fluentui/react-components";
import {
  getCategoryAccent,
  getSeverityLabel,
  isCommentOnlyCard,
  resolveCardVisualState,
} from "./ResultSuggestionCard.helpers";
import type { CardVisualState, CategoryAccent } from "./ResultSuggestionCard.helpers";
import type {
  ResultSuggestionCardClasses,
  ResultSuggestionCardViewModel,
} from "./ResultSuggestionCard.types";
import type { ResultsPanelCardState } from "../../ResultsPanelStore";

const useResultSuggestionCardStyles = makeStyles({
  root: {
    listStyle: "none",
    margin: 0,
    marginBottom: tokens.spacingVerticalM,
    padding: 0,
  },
  card: {
    position: "relative",
    paddingLeft: tokens.spacingHorizontalL,
    paddingRight: tokens.spacingHorizontalM,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
    rowGap: tokens.spacingVerticalS,
    overflow: "hidden",
    transitionProperty: "background-color, opacity, border-color",
    transitionDuration: tokens.durationNormal,
    transitionTimingFunction: tokens.curveEasyEase,
  },
  cardStatePending: {},
  cardStateAccepted: {
    backgroundColor: tokens.colorPaletteGreenBackground2,
    opacity: 0.78,
  },
  cardStateRejected: {
    backgroundColor: tokens.colorNeutralBackground3,
    opacity: 0.7,
  },
  cardStateFailed: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    borderTopColor: tokens.colorPaletteRedBorder1,
    borderRightColor: tokens.colorPaletteRedBorder1,
    borderBottomColor: tokens.colorPaletteRedBorder1,
    borderLeftColor: tokens.colorPaletteRedBorder1,
  },
  cardStateNotFound: {
    backgroundColor: tokens.colorPaletteDarkOrangeBackground2,
    borderTopColor: tokens.colorPaletteDarkOrangeBorder1,
    borderRightColor: tokens.colorPaletteDarkOrangeBorder1,
    borderBottomColor: tokens.colorPaletteDarkOrangeBorder1,
    borderLeftColor: tokens.colorPaletteDarkOrangeBorder1,
  },
  stripeBase: {
    position: "absolute",
    top: 0,
    bottom: 0,
    left: 0,
    width: "4px",
  },
  stripePending: {
    backgroundColor: tokens.colorBrandStroke1,
  },
  stripeAccepted: {
    backgroundColor: tokens.colorPaletteGreenBorder2,
  },
  stripeRejected: {
    backgroundColor: tokens.colorNeutralStroke1,
  },
  stripeFailed: {
    backgroundColor: tokens.colorPaletteRedBorder2,
  },
  stripeNotFound: {
    backgroundColor: tokens.colorPaletteDarkOrangeBorder2,
  },
  badgeRow: {
    display: "flex",
    flexWrap: "wrap",
    columnGap: tokens.spacingHorizontalS,
    rowGap: tokens.spacingVerticalXS,
    alignItems: "center",
  },
  categoryPillBase: {
    display: "inline-flex",
    alignItems: "center",
    height: "20px",
    paddingLeft: tokens.spacingHorizontalSNudge,
    paddingRight: tokens.spacingHorizontalSNudge,
    borderTopLeftRadius: tokens.borderRadiusCircular,
    borderTopRightRadius: tokens.borderRadiusCircular,
    borderBottomLeftRadius: tokens.borderRadiusCircular,
    borderBottomRightRadius: tokens.borderRadiusCircular,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase200,
    letterSpacing: "0.01em",
    borderTopWidth: "1px",
    borderRightWidth: "1px",
    borderBottomWidth: "1px",
    borderLeftWidth: "1px",
    borderTopStyle: "solid",
    borderRightStyle: "solid",
    borderBottomStyle: "solid",
    borderLeftStyle: "solid",
  },
  categoryGrammar: {
    color: tokens.colorPaletteCornflowerForeground2,
    backgroundColor: tokens.colorPaletteCornflowerBackground2,
    borderTopColor: tokens.colorPaletteCornflowerBorderActive,
    borderRightColor: tokens.colorPaletteCornflowerBorderActive,
    borderBottomColor: tokens.colorPaletteCornflowerBorderActive,
    borderLeftColor: tokens.colorPaletteCornflowerBorderActive,
  },
  categorySpelling: {
    color: tokens.colorPaletteLavenderForeground2,
    backgroundColor: tokens.colorPaletteLavenderBackground2,
    borderTopColor: tokens.colorPaletteLavenderBorderActive,
    borderRightColor: tokens.colorPaletteLavenderBorderActive,
    borderBottomColor: tokens.colorPaletteLavenderBorderActive,
    borderLeftColor: tokens.colorPaletteLavenderBorderActive,
  },
  categoryPunctuation: {
    color: tokens.colorPaletteLightTealForeground2,
    backgroundColor: tokens.colorPaletteLightTealBackground2,
    borderTopColor: tokens.colorPaletteLightTealBorderActive,
    borderRightColor: tokens.colorPaletteLightTealBorderActive,
    borderBottomColor: tokens.colorPaletteLightTealBorderActive,
    borderLeftColor: tokens.colorPaletteLightTealBorderActive,
  },
  categoryStyle: {
    color: tokens.colorPalettePeachForeground2,
    backgroundColor: tokens.colorPalettePeachBackground2,
    borderTopColor: tokens.colorPalettePeachBorderActive,
    borderRightColor: tokens.colorPalettePeachBorderActive,
    borderBottomColor: tokens.colorPalettePeachBorderActive,
    borderLeftColor: tokens.colorPalettePeachBorderActive,
  },
  categoryNeutral: {
    color: tokens.colorPaletteMinkForeground2,
    backgroundColor: tokens.colorPaletteMinkBackground2,
    borderTopColor: tokens.colorPaletteMinkBorderActive,
    borderRightColor: tokens.colorPaletteMinkBorderActive,
    borderBottomColor: tokens.colorPaletteMinkBorderActive,
    borderLeftColor: tokens.colorPaletteMinkBorderActive,
  },
  severityIndicator: {
    display: "inline-flex",
    alignItems: "center",
    columnGap: tokens.spacingHorizontalXXS,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
  },
  severityDotBase: {
    width: "6px",
    height: "6px",
    borderTopLeftRadius: "50%",
    borderTopRightRadius: "50%",
    borderBottomLeftRadius: "50%",
    borderBottomRightRadius: "50%",
    flexShrink: 0,
  },
  severityDotHigh: {
    backgroundColor: tokens.colorPaletteRedBackground3,
  },
  severityDotMedium: {
    backgroundColor: tokens.colorPaletteMarigoldBackground3,
  },
  severityDotLow: {
    backgroundColor: tokens.colorNeutralStroke1,
  },
  severityLabel: {
    fontWeight: tokens.fontWeightMedium,
    position: "relative",
    bottom: "2px",
    textTransform: "lowercase",
  },
  statusIconBase: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    marginLeft: "auto",
    fontSize: "18px",
    lineHeight: 1,
  },
  statusIconAccepted: {
    color: tokens.colorPaletteGreenForeground1,
  },
  statusIconRejected: {
    color: tokens.colorNeutralForeground3,
  },
  statusIconFailed: {
    color: tokens.colorPaletteRedForeground1,
  },
  statusIconNotFound: {
    color: tokens.colorPaletteDarkOrangeForeground1,
  },
  diff: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalXXS,
    paddingTop: tokens.spacingVerticalXS,
  },
  original: {
    color: tokens.colorPaletteRedForeground1,
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
  feedbackTextarea: {
    width: "100%",
  },
});

type Styles = ReturnType<typeof useResultSuggestionCardStyles>;

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

function resolveCardStateClass(styles: Styles, state: CardVisualState): string {
  if (state === "accepted") return styles.cardStateAccepted;
  if (state === "rejected") return styles.cardStateRejected;
  if (state === "failed") return styles.cardStateFailed;
  if (state === "not-found") return styles.cardStateNotFound;
  return styles.cardStatePending;
}

function resolveStripeClass(styles: Styles, state: CardVisualState): string {
  if (state === "accepted") return styles.stripeAccepted;
  if (state === "rejected") return styles.stripeRejected;
  if (state === "failed") return styles.stripeFailed;
  if (state === "not-found") return styles.stripeNotFound;
  return styles.stripePending;
}

function resolveCategoryClass(styles: Styles, accent: CategoryAccent): string {
  if (accent === "grammar") return styles.categoryGrammar;
  if (accent === "spelling") return styles.categorySpelling;
  if (accent === "punctuation") return styles.categoryPunctuation;
  if (accent === "style") return styles.categoryStyle;
  return styles.categoryNeutral;
}

function resolveSeverityDotClass(
  styles: Styles,
  severity: ResultsPanelCardState["suggestion"]["severity"]
): string {
  if (severity === "high") return styles.severityDotHigh;
  if (severity === "medium") return styles.severityDotMedium;
  return styles.severityDotLow;
}

function resolveStatusIconClass(styles: Styles, state: CardVisualState): string {
  if (state === "accepted") return styles.statusIconAccepted;
  if (state === "rejected") return styles.statusIconRejected;
  if (state === "failed") return styles.statusIconFailed;
  if (state === "not-found") return styles.statusIconNotFound;
  return "";
}
