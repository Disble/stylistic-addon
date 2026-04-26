import type { InsertionResult } from "../DocumentApplication.types";
import type { Suggestion } from "../suggestion/Suggestion.types";

/** Aggregate outcome of a completed analysis pipeline run. */
export interface PipelineResult {
  /** All unique suggestions produced by the analysis. */
  suggestions: Suggestion[];

  /** Insertion outcome. */
  result: InsertionResult;

  /** Error messages from chunks that failed analysis. */
  chunkErrors: string[];

  /** Whether the analysis was scoped to a text selection. */
  isSelection: boolean;

  /** Whether the pipeline was aborted before completion. */
  aborted: boolean;

  /** Human-readable reason for abortion. */
  abortReason?: string;
}
