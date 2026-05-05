import { makeStyles, tokens } from "@fluentui/react-components";
import {
  acceptResultsPanelSuggestion,
  computeResultsPanelChipCounts,
  navigateResultsPanelSuggestion,
  rejectResultsPanelSuggestion,
  selectResultsPanelVisibleCards,
  setResultsPanelFeedbackComment,
  setResultsPanelFilter,
  toggleResultsPanelFeedback,
  useResultsPanelStore,
} from "../../ResultsPanelStore";
import { useTaskpaneShellStore } from "../../TaskpaneShellStore";
import type { ResultsPanelClasses } from "./ResultsPanel.types";

const useResultsPanelStyles = makeStyles({
  root: {
    flex: "1 1 auto",
    minHeight: 0,
    overflowY: "auto",
    paddingBottom: tokens.spacingVerticalS,
  },
  list: {
    listStyle: "none",
    margin: 0,
    paddingTop: tokens.spacingVerticalS,
    paddingRight: 0,
    paddingBottom: 0,
    paddingLeft: 0,
  },
  empty: {
    paddingTop: tokens.spacingVerticalL,
    paddingBottom: tokens.spacingVerticalL,
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
  skeletonList: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalS,
    paddingTop: tokens.spacingVerticalS,
  },
  skeletonItem: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalXS,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    borderTopWidth: "1px",
    borderRightWidth: "1px",
    borderBottomWidth: "1px",
    borderLeftWidth: "1px",
    borderTopStyle: "solid",
    borderRightStyle: "solid",
    borderBottomStyle: "solid",
    borderLeftStyle: "solid",
    borderTopColor: tokens.colorNeutralStroke2,
    borderRightColor: tokens.colorNeutralStroke2,
    borderBottomColor: tokens.colorNeutralStroke2,
    borderLeftColor: tokens.colorNeutralStroke2,
    borderTopLeftRadius: tokens.borderRadiusMedium,
    borderTopRightRadius: tokens.borderRadiusMedium,
    borderBottomLeftRadius: tokens.borderRadiusMedium,
    borderBottomRightRadius: tokens.borderRadiusMedium,
  },
});

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
