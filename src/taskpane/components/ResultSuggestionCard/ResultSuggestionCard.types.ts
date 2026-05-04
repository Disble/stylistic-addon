import type { ResultsPanelCardState } from "../../ResultsPanelStore";

export type ResultSuggestionCardProps = Readonly<{
  card: ResultsPanelCardState;
  onAccept: (cardId: string) => Promise<void>;
  onFeedbackCommentChange: (cardId: string, feedbackComment: string) => void;
  onNavigate: (cardId: string) => Promise<void>;
  onReject: (cardId: string) => Promise<void>;
  onToggleFeedback: (cardId: string) => void;
}>;
