import * as React from "react";
import { ResultSuggestionCard } from "../ResultSuggestionCard";
import { useResultsPanel } from "./useResultsPanel";

/** Renders the results panel anchors consumed by the legacy DOM renderer. */
export function ResultsPanel(): React.JSX.Element {
  const {
    acceptSuggestion,
    cards,
    navigateToSuggestion,
    rejectSuggestion,
    setFeedbackComment,
    summaryText,
    toggleFeedback,
    visible,
  } = useResultsPanel();

  return (
    <div id="results-panel" className="results-panel" style={{ display: visible ? "block" : "none" }}>
      <h2 className="results-title">Resultados</h2>
      <div id="results-summary" className="results-summary">
        {summaryText}
      </div>
      <ul id="results-list" className="results-list">
        {cards.map((card) => (
          <ResultSuggestionCard
            card={card}
            key={card.suggestion.id}
            onAccept={acceptSuggestion}
            onFeedbackCommentChange={setFeedbackComment}
            onNavigate={navigateToSuggestion}
            onReject={rejectSuggestion}
            onToggleFeedback={toggleFeedback}
          />
        ))}
      </ul>
    </div>
  );
}
