import type { ApplySuggestionsResult } from "../DocumentApplication.types";
import type { Suggestion } from "../suggestion/Suggestion.types";
import type { PipelineState } from "./PipelineStateMachine.types";

/** Pipeline phase reported by document-application progress callbacks. */
export type AnalysisPhase = "applying";

/** Progress callback emitted by adapter-level suggestion application. */
export type ProgressCallback = (
  phase: AnalysisPhase,
  current: number,
  total: number,
  message: string
) => void;

/**
 * Observer for pipeline lifecycle events.
 * All methods are optional — implement only what you need.
 */
export interface PipelineObserver {
  /** Called when a pipeline phase begins. */
  onPhaseStart?(phase: PipelineState, message: string): void;

  /** Called periodically during a phase to report incremental progress. */
  onProgress?(current: number, total: number, message: string): void;

  /** Called when a pipeline phase completes successfully. */
  onPhaseComplete?(phase: PipelineState): void;

  /** Called when the pipeline encounters a recoverable or fatal error. */
  onError?(phase: PipelineState, error: Error | string): void;

  /** Called when the pipeline completes (success or partial success). */
  onComplete?(
    suggestions: Suggestion[],
    result: ApplySuggestionsResult,
    chunkErrors: string[],
    isSelection: boolean
  ): void;

  /** Called when the pipeline is aborted before reaching `done`. */
  onAbort?(reason: string): void;
}
