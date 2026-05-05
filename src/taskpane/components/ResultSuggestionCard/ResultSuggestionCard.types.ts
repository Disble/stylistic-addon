import type { ResultsPanelCardState } from "../../ResultsPanelStore";
import type { CardVisualState, CategoryAccent } from "./ResultSuggestionCard.helpers";

export type ResultSuggestionCardProps = Readonly<{
  card: ResultsPanelCardState;
  onAccept: (cardId: string) => Promise<void>;
  onFeedbackCommentChange: (cardId: string, feedbackComment: string) => void;
  onNavigate: (cardId: string) => Promise<void>;
  onReject: (cardId: string) => Promise<void>;
  onToggleFeedback: (cardId: string) => void;
}>;

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
