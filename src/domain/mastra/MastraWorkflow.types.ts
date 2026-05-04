import type {
  Suggestion,
  SuggestionSeverity,
  SuggestionType,
} from "../suggestion/Suggestion.types";

/** Input data sent to the Mastra stylistic workflow for each chunk. */
export interface WorkflowInput {
  /** Text to analyze. */
  text: string;

  /** Genre identifier matching the backend enum. */
  genero: "narrativa-literaria" | "ensayo-academico" | "periodismo-cultural" | "general";

  /** Author slug in kebab-case used to load the author profile. */
  autorSlug: string;
}

/** Raw suggestion shape as returned by the Mastra workflow. */
export interface WorkflowSuggestion {
  /** Paragraph-level context fragment used to locate the suggestion. */
  context: string;

  /** Exact substring of `context` targeted by the suggestion. */
  anchor: string;

  /** Replacement text. Absent when `type` is `"comment-only"`. */
  suggestedText?: string;

  /** Human-readable justification. */
  justification: string;

  /** Editorial category label. */
  category: string;

  /** How critical the suggestion is. */
  severity: SuggestionSeverity;

  /** Suggestion kind as declared by the backend. */
  type?: SuggestionType;
}

/** Expected output from the Mastra editorial workflow on success. */
export interface WorkflowOutput {
  /** Array of editorial suggestions for the analyzed text. */
  suggestions: WorkflowSuggestion[];

  /** Optional warnings from the backend. */
  warnings?: string[];
}

/** Workflow statuses observed while polling a chunk analysis run. */
export type ChunkAnalysisStatus =
  | "running"
  | "success"
  | "failed"
  | "tripwire"
  | "suspended"
  | "waiting"
  | "pending"
  | "canceled"
  | "bailed"
  | "paused";

/** Result of submitting a chunk for asynchronous workflow execution. */
export interface ChunkSubmitResult {
  /** Zero-based index of the chunk this submission belongs to. */
  chunkIndex: number;

  /** Workflow run identifier used for later polling. */
  runId?: string;

  /** Error message when submission could not be confirmed. */
  error?: string;
}

/** Result of polling an asynchronous workflow run for a chunk. */
export interface ChunkPollResult {
  /** Zero-based index of the chunk this poll belongs to. */
  chunkIndex: number;

  /** Workflow run identifier being polled. */
  runId: string;

  /** Current lifecycle status of the workflow run. */
  status: ChunkAnalysisStatus;

  /** Suggestions returned when the run completed successfully. */
  suggestions: Suggestion[];

  /** Error message for terminal failed states. */
  error?: string;
}

/** Outcome of analyzing a single text chunk via the Mastra workflow. */
export interface ChunkResult {
  /** Zero-based index of the chunk this result corresponds to. */
  chunkIndex: number;

  /** Suggestions returned by the workflow for this chunk. */
  suggestions: Suggestion[];

  /** Error message if the chunk analysis failed. */
  error?: string;
}
