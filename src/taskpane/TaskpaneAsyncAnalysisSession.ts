import { create } from "zustand";
import { INITIAL_TASKPANE_ASYNC_ANALYSIS_SESSION_STATE } from "./TaskpaneAsyncAnalysisSession.constants";
import type {
  TaskpaneAsyncAnalysisSessionSnapshot,
  TaskpaneAsyncAnalysisSessionState,
} from "./TaskpaneAsyncAnalysisSession.types";

/** Zustand store for the active async start/poll session shown in the taskpane. */
export const useTaskpaneAsyncAnalysisSessionStore = create<TaskpaneAsyncAnalysisSessionState>()(
  () => INITIAL_TASKPANE_ASYNC_ANALYSIS_SESSION_STATE
);

/** Resolves the store phase from the latest analysis snapshot. */
function resolveAsyncAnalysisPhase(
  snapshot: TaskpaneAsyncAnalysisSessionSnapshot
): TaskpaneAsyncAnalysisSessionState["phase"] {
  if (snapshot.activeAnalysisId !== null) {
    return "polling";
  }

  if (snapshot.retryableAnalysisId !== null) {
    return "retryable-failure";
  }

  return "idle";
}

/** Returns the current async analysis session snapshot. */
export function getTaskpaneAsyncAnalysisSessionState(): TaskpaneAsyncAnalysisSessionState {
  return useTaskpaneAsyncAnalysisSessionStore.getState();
}

/** Marks the beginning of a new start/poll cycle for the active session. */
export function startTaskpaneAsyncAnalysisSession(): void {
  useTaskpaneAsyncAnalysisSessionStore.setState({
    phase: "submitting",
    isSelection: false,
    activeAnalysisId: null,
    retryableAnalysisId: null,
  });
}

/** Publishes the current analysis snapshot into the taskpane session store. */
export function setTaskpaneAsyncAnalysisSnapshot(
  snapshot: TaskpaneAsyncAnalysisSessionSnapshot
): void {
  useTaskpaneAsyncAnalysisSessionStore.setState({
    phase: resolveAsyncAnalysisPhase(snapshot),
    isSelection: snapshot.isSelection,
    activeAnalysisId: snapshot.activeAnalysisId,
    retryableAnalysisId: snapshot.retryableAnalysisId,
  });
}

/** Marks that the frontend is sending the backend cancellation request. */
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
