/**
 * Shared TypeScript interfaces for the Stylistic add-in.
 *
 * These contracts define the communication boundaries between:
 * - The Mastra backend (workflow input/output)
 * - The Word API layer (suggestions and insertion results)
 * - The UI orchestrator (progress reporting)
 * - The pipeline (state machine, commands, events)
 *
 * No runtime code — only type declarations. No dependency on Office.js,
 * @mastra/client-js, or the DOM, keeping all consumers independently testable.
 *
 * @module types
 */

// ---------------------------------------------------------------------------
// Suggestion & Insertion
// ---------------------------------------------------------------------------

/**
 * A single editorial suggestion, either received from the Mastra workflow
 * or prepared for insertion into the Word document.
 */
export interface Suggestion {
  /** Unique identifier assigned by the frontend (e.g., "chunk0-3"). */
  id: string;

  /** Exact text to locate in the document (case-sensitive search via Word API). */
  originalText: string;

  /** Replacement text that will appear as a tracked change. */
  suggestedText: string;

  /** Human-readable reason for the suggestion, shown in the results panel. */
  justification: string;

  /** Editorial category label (e.g., "Redundancia", "Muletilla"). */
  category: string;

  /** How critical the suggestion is. */
  severity: "high" | "medium" | "low";
}

/**
 * Result of attempting to insert suggestions as tracked changes in Word.
 */
export interface InsertionResult {
  /** Number of suggestions successfully applied as tracked changes. */
  successCount: number;

  /** Suggestions whose `originalText` was not found in the document. */
  failedSuggestions: Suggestion[];
}

// ---------------------------------------------------------------------------
// Text Chunking
// ---------------------------------------------------------------------------

/**
 * A chunk of document text prepared for sending to the Mastra workflow.
 * Created by the chunker when the document exceeds the maximum chunk size.
 */
export interface TextChunk {
  /** The text content of this chunk. */
  text: string;

  /** Zero-based index of this chunk within the full document. */
  index: number;

  /** Total number of chunks the document was split into. */
  total: number;

  /** Character offset where this chunk starts in the original document. */
  startOffset: number;
}

// ---------------------------------------------------------------------------
// Mastra Workflow Communication
// ---------------------------------------------------------------------------

/**
 * Input data sent to the Mastra stylistic workflow for each chunk.
 * Must match the workflow's `inputSchema` on the backend (`stylistic-workflow`).
 */
export interface WorkflowInput {
  /** Text to analyze. */
  text: string;

  /** Genre identifier matching the backend enum (e.g., "narrativa-literaria", "general"). */
  genero: "narrativa-literaria" | "ensayo-academico" | "periodismo-cultural" | "general";

  /** Author slug in kebab-case used to load the author profile from the workspace. */
  autorSlug: string;
}

/**
 * Raw suggestion shape as returned by the Mastra workflow.
 * Does not include `id` — the frontend assigns IDs after receiving the response.
 */
export interface WorkflowSuggestion {
  /** Exact substring from the input text (case-sensitive). */
  originalText: string;

  /** Replacement text. */
  suggestedText: string;

  /** Human-readable justification. */
  justification: string;

  /** Editorial category label. */
  category: string;

  /** How critical the suggestion is. */
  severity: "high" | "medium" | "low";
}

/**
 * Expected output from the Mastra editorial workflow on success.
 * Accessed via `result.result` after a workflow run completes.
 */
export interface WorkflowOutput {
  /** Array of editorial suggestions for the analyzed text. */
  suggestions: WorkflowSuggestion[];

  /** Optional warnings from the backend (e.g., "text too short for meaningful analysis"). */
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

// ---------------------------------------------------------------------------
// Progress Reporting
// ---------------------------------------------------------------------------

/**
 * Phases of the analysis pipeline, used for progress reporting in the UI.
 */
export type AnalysisPhase = "reading" | "connecting" | "analyzing" | "applying" | "done";

/**
 * Callback signature for reporting progress during multi-phase analysis.
 *
 * @param phase   - Current phase of the pipeline.
 * @param current - Current step within the phase (1-based).
 * @param total   - Total steps in the phase.
 * @param message - Human-readable status message for display.
 */
export type ProgressCallback = (
  phase: AnalysisPhase,
  current: number,
  total: number,
  message: string
) => void;

// ---------------------------------------------------------------------------
// Chunk Analysis Result
// ---------------------------------------------------------------------------

/**
 * Outcome of analyzing a single text chunk via the Mastra workflow.
 * Used to aggregate results and track partial failures across chunks.
 */
export interface ChunkResult {
  /** Zero-based index of the chunk this result corresponds to. */
  chunkIndex: number;

  /** Suggestions returned by the workflow for this chunk (empty on failure). */
  suggestions: Suggestion[];

  /** Error message if the chunk analysis failed, `undefined` on success. */
  error?: string;
}

// ---------------------------------------------------------------------------
// Change Strategy
// ---------------------------------------------------------------------------

/**
 * Type of tracked change operation, used by the Strategy pattern in the
 * Word adapter.
 *
 * All types are applied via OOXML markup (`<w:del>`, `<w:ins>`) with an
 * attached Word comment containing the justification.
 *
 * - `"insert"` — Text insertion only (`<w:ins>` markup).
 * - `"delete"` — Text deletion only (`<w:del>` markup).
 * - `"replace"` — Combined deletion + insertion (`<w:del>` + `<w:ins>`).
 */
export type ChangeType = "insert" | "delete" | "replace";

// ---------------------------------------------------------------------------
// Text Source
// ---------------------------------------------------------------------------

/**
 * Result of the text-source resolution step at the start of the analysis
 * pipeline. Encapsulates whether the text came from the user's active
 * selection or from the full document body.
 */
export interface TextSource {
  /** The plain text to analyze (selection or full document). */
  text: string;

  /**
   * `true` if the text was read from the user's current selection;
   * `false` if it was read from the document body (no active selection).
   */
  isSelection: boolean;
}

// ---------------------------------------------------------------------------
// Profile
// ---------------------------------------------------------------------------

/**
 * An analysis profile option shown in the UI dropdown.
 */
export interface Profile {
  /** Machine-readable identifier sent to the workflow. */
  id: string;

  /** Human-readable label displayed in the dropdown. */
  label: string;
}

// ---------------------------------------------------------------------------
// Pipeline State Machine
// ---------------------------------------------------------------------------

/**
 * All possible states of the analysis pipeline.
 * Used by `PipelineStateMachine` to enforce valid transitions.
 *
 * - `idle`       — No pipeline running; ready to start.
 * - `reading`    — Reading text from the document or selection.
 * - `connecting` — Verifying backend connectivity.
 * - `chunking`   — Splitting text at paragraph boundaries.
 * - `analyzing`  — Sending chunks to the Mastra workflow.
 * - `applying`   — Applying suggestions as tracked changes.
 * - `done`       — Pipeline completed successfully.
 * - `error`      — Pipeline aborted due to unrecoverable error.
 */
export type PipelineState =
  | "idle"
  | "reading"
  | "connecting"
  | "chunking"
  | "analyzing"
  | "applying"
  | "done"
  | "error";

// ---------------------------------------------------------------------------
// Command Pattern
// ---------------------------------------------------------------------------

/**
 * Result of executing a `DocumentCommand`.
 */
export interface CommandResult {
  /** Whether the command completed successfully. */
  success: boolean;

  /** The ID of the command that produced this result. */
  commandId: string;

  /** Error message if `success` is false. */
  error?: string;
}

/**
 * A reversible document operation (Command pattern).
 *
 * Commands encapsulate a single document mutation (e.g., applying one
 * suggestion as a tracked change). The `execute()` method performs the
 * operation. An `undo()` method can be added in a future iteration.
 */
export interface DocumentCommand {
  /** Stable identifier matching the source suggestion's id. */
  readonly id: string;

  /** Human-readable description for logging and UI. */
  readonly description: string;

  /** Executes the command against the Word document. */
  execute(): Promise<CommandResult>;
}

// ---------------------------------------------------------------------------
// Pipeline Result
// ---------------------------------------------------------------------------

/**
 * The aggregate outcome of a completed analysis pipeline run.
 * Produced by `ApplySuggestionsHandler` and consumed by `taskpane.ts`
 * to render the results panel.
 */
export interface PipelineResult {
  /** All unique suggestions produced by the analysis (after dedup + guard). */
  suggestions: Suggestion[];

  /** Insertion outcome (success count + failed suggestions). */
  result: InsertionResult;

  /** Error messages from chunks that failed analysis. */
  chunkErrors: string[];

  /** Whether the analysis was scoped to a text selection. */
  isSelection: boolean;

  /** Whether the pipeline was aborted before completion. */
  aborted: boolean;

  /** Human-readable reason for abortion (if `aborted` is true). */
  abortReason?: string;
}

/**
 * Visual state of a suggestion card in the taskpane after user action.
 * - "pending": The suggestion has been applied and awaits user accept/reject.
 * - "accepted": The user accepted the suggestion from the taskpane.
 * - "rejected": The user rejected the suggestion from the taskpane.
 * - "already-resolved": The suggestion was resolved via Word's native Review panel.
 */
export type SuggestionState = "pending" | "accepted" | "rejected" | "already-resolved"

/**
 * Result returned by `IDocumentPort.acceptSuggestion` and `IDocumentPort.rejectSuggestion`.
 */
export interface SuggestionActionResult {
  /** Final resolution status of the operation. */
  status: "accepted" | "rejected" | "already-resolved" | "not-found" | "error"
  /** Number of tracked changes (insert + delete) that were accepted or rejected. */
  trackedChangesAffected: number
  /** Whether the associated Stylistic comment was successfully deleted. */
  commentDeleted: boolean
  /** Human-readable error message when status is "error" or "not-found". */
  error?: string
}
