import type { AnalysisSnapshot } from "./Analysis.types";

/** Terminal analysis snapshot, or the reason polling stopped without one. */
export type AnalysisPollOutcome =
  | { kind: "finished"; snapshot: AnalysisSnapshot }
  | { kind: "canceled-locally" }
  | { kind: "poll-failed"; message: string };

/** Hooks a caller passes to observe or stop polling. */
export interface AnalysisPollOptions {
  /** Returns true once the user asked to stop waiting. */
  shouldCancel?: () => boolean;

  /** Called after every successful read of the analysis. */
  onPoll?: (snapshot: AnalysisSnapshot) => void;
}
