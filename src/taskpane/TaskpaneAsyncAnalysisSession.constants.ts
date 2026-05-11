import type { TaskpaneAsyncAnalysisSessionState } from "./TaskpaneAsyncAnalysisSession.types";

/** Initial in-memory async session state. Persisted only for the active taskpane lifetime. */
export const INITIAL_TASKPANE_ASYNC_ANALYSIS_SESSION_STATE: TaskpaneAsyncAnalysisSessionState = {
  phase: "idle",
  isSelection: false,
  activeRuns: [],
  retryableRuns: [],
};
