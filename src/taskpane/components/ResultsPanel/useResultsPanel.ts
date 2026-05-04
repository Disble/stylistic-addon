import { useSyncExternalStore } from "react";
import {
  acceptResultsPanelSuggestion,
  getResultsPanelState,
  navigateResultsPanelSuggestion,
  rejectResultsPanelSuggestion,
  setResultsPanelFeedbackComment,
  subscribeResultsPanelStore,
  toggleResultsPanelFeedback,
} from "../../ResultsPanelStore";

/** React hook for the results-panel store and interaction commands. */
export function useResultsPanel() {
  const state = useSyncExternalStore(subscribeResultsPanelStore, getResultsPanelState);

  return {
    ...state,
    acceptSuggestion: acceptResultsPanelSuggestion,
    navigateToSuggestion: navigateResultsPanelSuggestion,
    rejectSuggestion: rejectResultsPanelSuggestion,
    setFeedbackComment: setResultsPanelFeedbackComment,
    toggleFeedback: toggleResultsPanelFeedback,
  };
}
