import {
  acceptResultsPanelSuggestion,
  navigateResultsPanelSuggestion,
  rejectResultsPanelSuggestion,
  setResultsPanelFeedbackComment,
  toggleResultsPanelFeedback,
  useResultsPanelStore,
} from "../../ResultsPanelStore";

/** React hook for the results-panel store and interaction commands. */
export function useResultsPanel() {
  const state = useResultsPanelStore();

  return {
    ...state,
    acceptSuggestion: acceptResultsPanelSuggestion,
    navigateToSuggestion: navigateResultsPanelSuggestion,
    rejectSuggestion: rejectResultsPanelSuggestion,
    setFeedbackComment: setResultsPanelFeedbackComment,
    toggleFeedback: toggleResultsPanelFeedback,
  };
}
