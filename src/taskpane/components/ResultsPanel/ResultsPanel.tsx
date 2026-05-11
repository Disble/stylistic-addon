import * as React from "react";
import { Caption1 } from "@fluentui/react-components";
import { ResultSuggestionCard } from "../ResultSuggestionCard";
import { ResultsSummaryChips } from "../ResultsSummaryChips";
import { useResultsPanel } from "./ResultsPanel.hooks";

/** Renders the chips toolbar and the filtered suggestion list. */
export function ResultsPanel(): React.JSX.Element | null {
  const {
    acceptSuggestion,
    activeFilter,
    cards,
    classes,
    counts,
    navigateToSuggestion,
    rejectSuggestion,
    setFeedbackComment,
    setFilter,
    summaryText,
    toggleFeedback,
    visible,
    visibleCards,
  } = useResultsPanel();

  if (!visible) {
    return null;
  }

  const hasCards = cards.length > 0;
  const filterMatches = visibleCards.length > 0;

  return (
    <div className={classes.root} data-testid="results-panel">
      {hasCards ? (
        <ResultsSummaryChips
          activeFilter={activeFilter}
          counts={counts}
          onFilterChange={setFilter}
          summaryText={summaryText}
        />
      ) : null}

      {filterMatches ? (
        <ul className={classes.list} data-testid="results-list">
          {visibleCards.map((card) => (
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
      ) : (
        <Caption1 className={classes.empty} data-testid="results-empty">
          {hasCards ? "Sin sugerencias en este filtro." : "No hay sugerencias para mostrar."}
        </Caption1>
      )}
    </div>
  );
}
