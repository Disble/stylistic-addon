import type { TaskpaneAsyncAnalysisSessionState } from "./TaskpaneAsyncAnalysisSession.types";
import type { AnalysisProfileId } from "../domain/Profile.types";

/** Visual intent of the transient taskpane status bar. */
export type TaskpaneStatusType = "success" | "error";

/** Retry path exposed by the explicit analysis-error surface. */
export type TaskpaneAnalysisRetryKind = "full-retry" | "retry-query";

/** Persistent-in-session analysis failure state shown outside the status bar. */
export type TaskpaneAnalysisErrorState = Readonly<{
  message: string;
  retryKind: TaskpaneAnalysisRetryKind;
  visible: boolean;
}>;

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
  asyncSession: TaskpaneAsyncAnalysisSessionState;
}>;

/** Root Zustand state for taskpane shell controls. */
export type TaskpaneShellState = Readonly<{
  analysisError: TaskpaneAnalysisErrorState;
  cleanupVisible: boolean;
  disableTrackChangesCtaVisible: boolean;
  isCleanupLoading: boolean;
  isDisableTrackChangesLoading: boolean;
  isAnalyzeLoading: boolean;
  selectedGenero: AnalysisProfileId;
  progress: TaskpaneShellProgress;
  status: TaskpaneShellStatus;
}>;
