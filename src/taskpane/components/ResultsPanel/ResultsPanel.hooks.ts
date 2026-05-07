import {
  acceptResultsPanelSuggestion,
  navigateResultsPanelSuggestion,
  rejectResultsPanelSuggestion,
  setResultsPanelFeedbackComment,
  setResultsPanelFilter,
  toggleResultsPanelFeedback,
  useResultsPanelStore,
} from "../../ResultsPanelStore";
import {
  computeResultsPanelChipCounts,
  selectResultsPanelVisibleCards,
} from "../../ResultsPanelFilters";
import { useTaskpaneShellStore } from "../../TaskpaneShellStore";
import { useResultsPanelStyles } from "./ResultsPanel.styles";
import type { ResultsPanelClasses } from "./ResultsPanel.types";

/** Returns Griffel classes for the results panel layout. */
function useResultsPanelClasses(): ResultsPanelClasses {
  const styles = useResultsPanelStyles();
  return {
    root: styles.root,
    list: styles.list,
    empty: styles.empty,
    skeletonList: styles.skeletonList,
    skeletonItem: styles.skeletonItem,
  };
}

/** React hook for the results-panel store, derived selectors, and interaction commands. */
export function useResultsPanel() {
  const state = useResultsPanelStore();
  const isAnalyzeLoading = useTaskpaneShellStore((shell) => shell.isAnalyzeLoading);
  const classes = useResultsPanelClasses();

  const counts = computeResultsPanelChipCounts(state.cards);
  const visibleCards = selectResultsPanelVisibleCards(state.cards, state.activeFilter);
  const showSkeleton = isAnalyzeLoading && state.cards.length === 0;

  return {
    activeFilter: state.activeFilter,
    acceptSuggestion: acceptResultsPanelSuggestion,
    cards: state.cards,
    classes,
    counts,
    navigateToSuggestion: navigateResultsPanelSuggestion,
    rejectSuggestion: rejectResultsPanelSuggestion,
    setFeedbackComment: setResultsPanelFeedbackComment,
    setFilter: setResultsPanelFilter,
    showSkeleton,
    summaryText: state.summaryText,
    toggleFeedback: toggleResultsPanelFeedback,
    visible: state.visible,
    visibleCards,
  };
}
