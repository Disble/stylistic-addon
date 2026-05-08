import { create } from "zustand";
import type { AnalysisProfileId } from "../domain/Profile.types";
import {
  HIDE_PROGRESS_DELAY_MS,
  INITIAL_TASKPANE_SHELL_STATE,
  STATUS_DISPLAY_MS,
} from "./TaskpaneShellStore.constants";
import type { TaskpaneShellState, TaskpaneStatusType } from "./TaskpaneShellStore.types";

let hideProgressTimeoutId: ReturnType<typeof setTimeout> | undefined;
let hideStatusTimeoutId: ReturnType<typeof setTimeout> | undefined;

/** Zustand store holding the reactive taskpane shell state. */
export const useTaskpaneShellStore = create<TaskpaneShellState>()(
  () => INITIAL_TASKPANE_SHELL_STATE
);

/** Returns the current immutable shell state snapshot. */
export function getTaskpaneShellState(): TaskpaneShellState {
  return useTaskpaneShellStore.getState();
}

/** Sets the loading state for the analyze CTA and related shell controls. */
export function setTaskpaneAnalyzeLoading(isAnalyzeLoading: boolean): void {
  useTaskpaneShellStore.setState({ isAnalyzeLoading });
}

/** Sets the selected analysis profile used by the composition root. */
export function setTaskpaneSelectedGenero(selectedGenero: AnalysisProfileId): void {
  useTaskpaneShellStore.setState({ selectedGenero });
}

/** Sets the cleanup CTA visibility. */
export function setTaskpaneCleanupVisible(cleanupVisible: boolean): void {
  useTaskpaneShellStore.setState({ cleanupVisible });
}

/** Sets the loading state of the cleanup CTA. */
export function setTaskpaneCleanupLoading(isCleanupLoading: boolean): void {
  useTaskpaneShellStore.setState({ isCleanupLoading });
}

/** Sets the Track Changes CTA visibility. */
export function setTaskpaneDisableTrackChangesCtaVisible(
  disableTrackChangesCtaVisible: boolean
): void {
  useTaskpaneShellStore.setState({ disableTrackChangesCtaVisible });
}

/** Sets the loading state of the Track Changes CTA. */
export function setTaskpaneDisableTrackChangesLoading(isDisableTrackChangesLoading: boolean): void {
  useTaskpaneShellStore.setState({ isDisableTrackChangesLoading });
}

/** Updates the visible progress state. */
export function updateTaskpaneProgress(current: number, total: number, message: string): void {
  clearTimeout(hideProgressTimeoutId);
  useTaskpaneShellStore.setState({
    progress: {
      current,
      total,
      message,
      visible: true,
    },
  });
}

/** Hides the progress area after the existing UX delay. */
export function hideTaskpaneProgress(): void {
  clearTimeout(hideProgressTimeoutId);
  hideProgressTimeoutId = setTimeout(() => {
    useTaskpaneShellStore.setState((state) => ({
      progress: {
        ...state.progress,
        visible: false,
      },
    }));
  }, HIDE_PROGRESS_DELAY_MS);
}

/** Shows a transient status-bar message and auto-hides it. */
export function showTaskpaneStatus(message: string, type: TaskpaneStatusType): void {
  clearTimeout(hideStatusTimeoutId);
  useTaskpaneShellStore.setState({
    status: {
      message,
      type,
      visible: true,
    },
  });

  hideStatusTimeoutId = setTimeout(() => {
    useTaskpaneShellStore.setState((state) => ({
      status: {
        ...state.status,
        visible: false,
      },
    }));
  }, STATUS_DISPLAY_MS);
}

/** Resets shell state and timers for deterministic tests. */
export function resetTaskpaneShellState(): void {
  clearTimeout(hideProgressTimeoutId);
  clearTimeout(hideStatusTimeoutId);
  hideProgressTimeoutId = undefined;
  hideStatusTimeoutId = undefined;
  useTaskpaneShellStore.setState(INITIAL_TASKPANE_SHELL_STATE, true);
}
