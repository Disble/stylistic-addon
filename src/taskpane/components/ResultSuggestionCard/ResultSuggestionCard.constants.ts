import type { CardVisualState } from "./ResultSuggestionCard.types";

/** Accessible label used for each rendered card status icon. */
export const STATUS_ICON_LABEL: Record<CardVisualState, string | null> = {
  pending: null,
  accepted: "Sugerencia aceptada",
  rejected: "Sugerencia rechazada",
  failed: "Sugerencia fallida",
  "not-found": "Sugerencia no encontrada en el documento",
};

/** Humanized severity label used by the inline severity indicator. */
export const RESULT_SUGGESTION_CARD_SEVERITY_LABEL = {
  high: "alta",
  medium: "media",
  low: "baja",
} as const;

/** Regex used to normalize accents before resolving category slugs. */
export const RESULT_SUGGESTION_CARD_DIACRITIC_PATTERN = /[̀-ͯ]/g;
