import { create } from "zustand";
import { INITIAL_TASKPANE_ASYNC_ANALYSIS_SESSION_STATE } from "./TaskpaneAsyncAnalysisSession.constants";
import type {
  TaskpaneAsyncAnalysisSessionSnapshot,
  TaskpaneAsyncAnalysisSessionState,
} from "./TaskpaneAsyncAnalysisSession.types";

/** Zustand store for the active async submit/poll session shown in the taskpane. */
export const useTaskpaneAsyncAnalysisSessionStore = create<TaskpaneAsyncAnalysisSessionState>()(
  () => INITIAL_TASKPANE_ASYNC_ANALYSIS_SESSION_STATE
);

/** Returns the current async analysis session snapshot. */
export function getTaskpaneAsyncAnalysisSessionState(): TaskpaneAsyncAnalysisSessionState {
  return useTaskpaneAsyncAnalysisSessionStore.getState();
}

/** Marks the beginning of a new submit/poll cycle for the active session. */
export function startTaskpaneAsyncAnalysisSession(): void {
  useTaskpaneAsyncAnalysisSessionStore.setState({
    phase: "submitting",
    isSelection: false,
    activeRuns: [],
    retryableRuns: [],
  });
}

/** Publishes the current backend-run snapshot into the taskpane session store. */
export function setTaskpaneAsyncAnalysisSnapshot(
  snapshot: TaskpaneAsyncAnalysisSessionSnapshot
): void {
  const hasActiveRuns = snapshot.activeRuns.length > 0;
  const hasRetryableRuns = snapshot.retryableRuns.length > 0;

  useTaskpaneAsyncAnalysisSessionStore.setState({
    phase: hasActiveRuns ? "polling" : hasRetryableRuns ? "retryable-failure" : "idle",
    isSelection: snapshot.isSelection,
    activeRuns: snapshot.activeRuns,
    retryableRuns: snapshot.retryableRuns,
  });
}

/** Marks that the frontend is sending backend cancellation requests. */
export function setTaskpaneAsyncAnalysisCanceling(): void {
  useTaskpaneAsyncAnalysisSessionStore.setState((state) => ({
    ...state,
    phase: "canceling",
  }));
}

/** Clears the async analysis session after success, terminal failure, or explicit cancel. */
export function clearTaskpaneAsyncAnalysisSession(): void {
  useTaskpaneAsyncAnalysisSessionStore.setState(
    INITIAL_TASKPANE_ASYNC_ANALYSIS_SESSION_STATE,
    true
  );
}
