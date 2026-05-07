import type { ResultsPanelState } from "./ResultsPanelStore.types";

/** Initial hidden state for the taskpane results panel store. */
export const INITIAL_RESULTS_PANEL_STATE: ResultsPanelState = {
  activeFilter: "all",
  cards: [],
  summaryText: "",
  visible: false,
};
