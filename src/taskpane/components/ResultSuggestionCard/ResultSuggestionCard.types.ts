import type { ResultsPanelCardState } from "../../ResultsPanelStore.types";
import type { useResultSuggestionCardStyles } from "./ResultSuggestionCard.styles";

/** Visual state used to tint one rendered suggestion card. */
export type CardVisualState = "pending" | "accepted" | "rejected" | "failed" | "not-found";

/** Accent slug used to color the category pill. */
export type CategoryAccent = "grammar" | "spelling" | "punctuation" | "style" | "neutral";

/** Griffel style hook return type consumed by card helper resolvers. */
export type ResultSuggestionCardStyles = ReturnType<typeof useResultSuggestionCardStyles>;

/** Props required to render one suggestion card in the results panel. */
export type ResultSuggestionCardProps = Readonly<{
  card: ResultsPanelCardState;
  onAccept: (cardId: string) => Promise<void>;
  onFeedbackCommentChange: (cardId: string, feedbackComment: string) => void;
  onNavigate: (cardId: string) => Promise<void>;
  onReject: (cardId: string) => Promise<void>;
  onToggleFeedback: (cardId: string) => void;
}>;

/** Griffel class slots consumed by the result suggestion card. */
export type ResultSuggestionCardClasses = Readonly<{
  root: string;
  card: string;
  stateStripe: string;
  badgeRow: string;
  categoryPill: string;
  severityIndicator: string;
  severityDot: string;
  severityLabel: string;
  statusIcon: string;
  diff: string;
  original: string;
  suggested: string;
  justification: string;
  clickableArea: string;
  failureLabel: string;
  failureDetail: string;
  note: string;
  footer: string;
  actions: string;
  feedbackTextarea: string;
}>;

/** View model consumed by the result suggestion card component. */
export type ResultSuggestionCardViewModel = Readonly<{
  classes: ResultSuggestionCardClasses;
  isCommentOnly: boolean;
  cardVisualState: CardVisualState;
  categoryAccent: CategoryAccent;
  severityLabel: string;
  acceptLabel: string;
  rejectLabel: string;
  acceptAriaLabel: string;
  rejectAriaLabel: string;
  feedbackToggleAriaLabel: string;
}>;
