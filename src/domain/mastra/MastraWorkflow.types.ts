import type {
  Suggestion,
  SuggestionSeverity,
  SuggestionType,
} from "../suggestion/Suggestion.types";

/** Editorial genres supported by the analysis workflow. */
export type WorkflowGenero =
  | "narrativa-literaria"
  | "ensayo-academico"
  | "periodismo-cultural"
  | "general";

/** Opaque document-level processing preferences forwarded to the backend. */
export type WorkflowProcessingConfig = Record<string, unknown>;

/** Document-scoped metadata submitted alongside each chunk analysis. */
export interface WorkflowSubmitContext {
  /** Stable document UUID generated and persisted by the add-in. */
  documentUuid: string;

  /** Editorial genre for analysis style. */
  genero?: WorkflowGenero;

  /** Optional document title shown or stored by the backend. */
  title?: string;

  /** Optional document-level processing preferences. */
  processingConfig?: WorkflowProcessingConfig;
}

/** Input data sent to the Mastra stylistic workflow for each chunk. */
export interface WorkflowInput extends WorkflowSubmitContext {
  /** Text to analyze. */
  text: string;

  /** Genre identifier matching the backend enum. */
  genero?: WorkflowGenero;
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
  | "paused"
  | "retryable-failure";

/** Describes which layer produced the current chunk poll outcome. */
export type ChunkPollOrigin = "backend" | "frontend-retryable" | "frontend-terminal";

/** Stable run reference reused across polling, cancellation, and retry-query flows. */
export interface ChunkRunReference {
  /** Zero-based index of the chunk associated with the run. */
  chunkIndex: number;

  /** Workflow run identifier issued by Mastra. */
  runId: string;
}

/** Snapshot of the submit/poll session exposed to the taskpane shell. */
export interface AnalysisRunSessionSnapshot {
  /** Whether the originating analysis targeted the current selection. */
  isSelection: boolean;

  /** Active backend runs still being polled by the frontend. */
  activeRuns: ChunkRunReference[];

  /** Runs whose latest failure is retryable from the frontend only. */
  retryableRuns: ChunkRunReference[];
}

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

  /** Indicates whether the state came from Mastra or from a frontend retry decision. */
  origin: ChunkPollOrigin;

  /** Suggestions returned when the run completed successfully. */
  suggestions: Suggestion[];

  /** Error message for terminal failed states. */
  error?: string;
}

/** Result of requesting cancellation for one active chunk run. */
export interface ChunkCancelResult extends ChunkRunReference {
  /** Whether the backend acknowledged the cancellation request. */
  canceled: boolean;

  /** Error message when the frontend could not confirm cancellation. */
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
