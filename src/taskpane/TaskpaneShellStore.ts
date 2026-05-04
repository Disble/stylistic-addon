import { useSyncExternalStore } from "react";

export type TaskpaneStatusType = "success" | "error";

export type TaskpaneShellStatus = Readonly<{
  message: string;
  type: TaskpaneStatusType;
  visible: boolean;
}>;

export type TaskpaneShellProgress = Readonly<{
  current: number;
  total: number;
  message: string;
  visible: boolean;
}>;

export type TaskpaneShellState = Readonly<{
  cleanupVisible: boolean;
  disableTrackChangesCtaVisible: boolean;
  isCleanupLoading: boolean;
  isDisableTrackChangesLoading: boolean;
  isAnalyzeLoading: boolean;
  selectedGenero: string;
  progress: TaskpaneShellProgress;
  status: TaskpaneShellStatus;
}>;

const STATUS_DISPLAY_MS = 4000;
const HIDE_PROGRESS_DELAY_MS = 1000;

const INITIAL_STATE: TaskpaneShellState = {
  cleanupVisible: false,
  disableTrackChangesCtaVisible: false,
  isCleanupLoading: false,
  isDisableTrackChangesLoading: false,
  isAnalyzeLoading: false,
  selectedGenero: "narrativa-literaria",
  progress: {
    current: 0,
    total: 1,
    message: "",
    visible: false,
  },
  status: {
    message: "",
    type: "success",
    visible: false,
  },
};

let shellState = INITIAL_STATE;
let hideProgressTimeoutId: ReturnType<typeof setTimeout> | undefined;
let hideStatusTimeoutId: ReturnType<typeof setTimeout> | undefined;

const listeners = new Set<() => void>();

/** Subscribes React consumers to shell-state changes. */
export function subscribeTaskpaneShellStore(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** Returns the current immutable shell state snapshot. */
export function getTaskpaneShellState(): TaskpaneShellState {
  return shellState;
}

/** React hook for consuming the external taskpane shell store. */
export function useTaskpaneShellState(): TaskpaneShellState {
  return useSyncExternalStore(subscribeTaskpaneShellStore, getTaskpaneShellState);
}

/** Sets the loading state for the analyze CTA and related shell controls. */
export function setTaskpaneAnalyzeLoading(isAnalyzeLoading: boolean): void {
  updateTaskpaneShellState({ isAnalyzeLoading });
}

/** Sets the selected analysis profile used by the composition root. */
export function setTaskpaneSelectedGenero(selectedGenero: string): void {
  updateTaskpaneShellState({ selectedGenero });
}

/** Sets the cleanup CTA visibility. */
export function setTaskpaneCleanupVisible(cleanupVisible: boolean): void {
  updateTaskpaneShellState({ cleanupVisible });
}

/** Sets the loading state of the cleanup CTA. */
export function setTaskpaneCleanupLoading(isCleanupLoading: boolean): void {
  updateTaskpaneShellState({ isCleanupLoading });
}

/** Sets the Track Changes CTA visibility. */
export function setTaskpaneDisableTrackChangesCtaVisible(
  disableTrackChangesCtaVisible: boolean
): void {
  updateTaskpaneShellState({ disableTrackChangesCtaVisible });
}

/** Sets the loading state of the Track Changes CTA. */
export function setTaskpaneDisableTrackChangesLoading(isDisableTrackChangesLoading: boolean): void {
  updateTaskpaneShellState({ isDisableTrackChangesLoading });
}

/** Updates the visible progress state. */
export function updateTaskpaneProgress(current: number, total: number, message: string): void {
  clearTimeout(hideProgressTimeoutId);
  updateTaskpaneShellState({
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
    updateTaskpaneShellState({
      progress: {
        ...shellState.progress,
        visible: false,
      },
    });
  }, HIDE_PROGRESS_DELAY_MS);
}

/** Shows a transient status-bar message and auto-hides it. */
export function showTaskpaneStatus(message: string, type: TaskpaneStatusType): void {
  clearTimeout(hideStatusTimeoutId);
  updateTaskpaneShellState({
    status: {
      message,
      type,
      visible: true,
    },
  });

  hideStatusTimeoutId = setTimeout(() => {
    updateTaskpaneShellState({
      status: {
        ...shellState.status,
        visible: false,
      },
    });
  }, STATUS_DISPLAY_MS);
}

/** Resets shell state and timers for deterministic tests. */
export function resetTaskpaneShellState(): void {
  clearTimeout(hideProgressTimeoutId);
  clearTimeout(hideStatusTimeoutId);
  hideProgressTimeoutId = undefined;
  hideStatusTimeoutId = undefined;
  shellState = INITIAL_STATE;

  for (const listener of listeners) {
    listener();
  }
}

/** Applies a partial immutable state update and notifies all subscribers. */
function updateTaskpaneShellState(partialState: Partial<TaskpaneShellState>): void {
  shellState = {
    ...shellState,
    ...partialState,
  };

  for (const listener of listeners) {
    listener();
  }
}
