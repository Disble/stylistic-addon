import type { TaskpaneShellState } from "./TaskpaneShellStore.types";

/** Milliseconds a status message remains visible before auto-hide. */
export const STATUS_DISPLAY_MS = 4000;

/** Milliseconds the progress area waits before hiding after completion. */
export const HIDE_PROGRESS_DELAY_MS = 1000;

/** Initial shell presentation state for taskpane boot and test resets. */
export const INITIAL_TASKPANE_SHELL_STATE: TaskpaneShellState = {
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
