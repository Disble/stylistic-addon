import * as React from "react";
import { Caption1, Skeleton, SkeletonItem } from "@fluentui/react-components";
import { ResultSuggestionCard } from "../ResultSuggestionCard";
import { ResultsSummaryChips } from "../ResultsSummaryChips";
import { SKELETON_PLACEHOLDER_COUNT } from "./ResultsPanel.constants";
import { useResultsPanel } from "./ResultsPanel.hooks";

/** Renders the chips toolbar, filtered suggestion list, and skeleton loading state. */
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
    showSkeleton,
    summaryText,
    toggleFeedback,
    visible,
    visibleCards,
  } = useResultsPanel();

  if (showSkeleton) {
    return (
      <div className={classes.root} data-testid="results-panel-skeleton">
        <div className={classes.skeletonList}>
          {Array.from({ length: SKELETON_PLACEHOLDER_COUNT }, (_, index) => (
            <Skeleton key={index} className={classes.skeletonItem}>
              <SkeletonItem shape="rectangle" size={16} />
              <SkeletonItem shape="rectangle" size={16} />
              <SkeletonItem shape="rectangle" size={12} />
            </Skeleton>
          ))}
        </div>
      </div>
    );
  }

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
