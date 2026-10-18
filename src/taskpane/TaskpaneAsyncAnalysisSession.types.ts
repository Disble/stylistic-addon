import type { AnalysisSessionSnapshot } from "../domain/analysis/Analysis.types";

/** Explicit frontend-owned async analysis phases for the active taskpane session. */
export type TaskpaneAsyncAnalysisPhase =
  | "idle"
  | "submitting"
  | "polling"
  | "retryable-failure"
  | "canceling";

/** Reactive session model for cancel/retry-query UX in the taskpane. */
export type TaskpaneAsyncAnalysisSessionState = Readonly<{
  phase: TaskpaneAsyncAnalysisPhase;
  isSelection: boolean;
  activeAnalysisId: string | null;
  retryableAnalysisId: string | null;
}>;

/** Helper payload used when replacing the active session snapshot from pipeline state. */
export type TaskpaneAsyncAnalysisSessionSnapshot = AnalysisSessionSnapshot;
