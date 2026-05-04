/** Phases of the analysis pipeline, used for progress reporting in the UI. */
export type AnalysisPhase = "reading" | "connecting" | "analyzing" | "applying" | "done";

/** Callback signature for reporting progress during multi-phase analysis. */
export type ProgressCallback = (
  phase: AnalysisPhase,
  current: number,
  total: number,
  message: string
) => void;
