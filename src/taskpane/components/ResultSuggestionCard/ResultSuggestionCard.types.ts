import type { ResultsPanelCardState } from "../../ResultsPanelStore";
import type { SeverityBadgeColor } from "./ResultSuggestionCard.helpers";

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
  severityStripe: string;
  badgeRow: string;
  diff: string;
  diffArrow: string;
  original: string;
  suggested: string;
  justification: string;
  clickableArea: string;
  failureLabel: string;
  failureDetail: string;
  note: string;
  footer: string;
  actions: string;
  primaryAction: string;
  feedbackTextarea: string;
}>;

export type ResultSuggestionCardViewModel = Readonly<{
  classes: ResultSuggestionCardClasses;
  isCommentOnly: boolean;
  severityColor: SeverityBadgeColor;
  severityLabel: string;
  acceptLabel: string;
  rejectLabel: string;
  acceptAriaLabel: string;
  rejectAriaLabel: string;
  feedbackToggleAriaLabel: string;
}>;
