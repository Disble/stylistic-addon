import type { AnalysisProfileId } from "../domain/Profile.types";

/** Visual intent of the transient taskpane status bar. */
export type TaskpaneStatusType = "success" | "error";

/** Visible taskpane status-bar payload. */
export type TaskpaneShellStatus = Readonly<{
  message: string;
  type: TaskpaneStatusType;
  visible: boolean;
}>;

/** Visible taskpane pipeline-progress payload. */
export type TaskpaneShellProgress = Readonly<{
  current: number;
  total: number;
  message: string;
  visible: boolean;
}>;

/** Root Zustand state for taskpane shell controls. */
export type TaskpaneShellState = Readonly<{
  cleanupVisible: boolean;
  disableTrackChangesCtaVisible: boolean;
  isCleanupLoading: boolean;
  isDisableTrackChangesLoading: boolean;
  isAnalyzeLoading: boolean;
  selectedGenero: AnalysisProfileId;
  progress: TaskpaneShellProgress;
  status: TaskpaneShellStatus;
}>;
