import type { ApplySuggestionsResult } from "../domain/DocumentApplication.types";
import type { Suggestion } from "../domain/suggestion/Suggestion.types";

/** Terminal progress buckets for suggestions already applied to the document. */
export type AppliedSuggestionProgressState = "pending" | "resolved" | "needs-attention";

/** Mutable progress snapshot owned by the results panel during one render cycle. */
export interface SuggestionProgressSummaryModel {
  /** Total suggestions emitted by the analysis pipeline. */
  total: number;

  /** Suggestions successfully materialized in Word during the initial batch. */
  applied: number;

  /** Suggestions never found in the document during the initial batch. */
  notFound: number;

  /** Suggestions that failed to apply for reasons other than not-found. */
  failedOther: number;

  /** Chunk-level analysis errors preserved from the pipeline. */
  chunkErrors: number;

  /** Per-suggestion live state for suggestions that were initially applied. */
  appliedStates: Map<string, AppliedSuggestionProgressState>;
}

/** Input payload needed to create the live suggestion-progress summary model. */
export type SuggestionProgressSummaryInput = Readonly<{
  suggestions: Suggestion[];
  result: ApplySuggestionsResult;
  chunkErrors: string[];
}>;
