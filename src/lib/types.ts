/**
 * Shared TypeScript interfaces for the Stylistic add-in.
 *
 * These contracts define the communication boundaries between:
 * - The Mastra backend (workflow input/output)
 * - The Word API layer (suggestions and insertion results)
 * - The UI orchestrator (progress reporting)
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
 * Input data sent to the Mastra editorial workflow for each chunk.
 * Must match the workflow's `inputSchema` on the backend.
 */
export interface WorkflowInput {
  /** Text to analyze. */
  text: string;

  /** Analysis profile identifier (e.g., "general", "formal", "academic"). */
  profile: string;

  /** ISO 639-1 language code of the text (e.g., "es", "en"). */
  language: string;
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

// ---------------------------------------------------------------------------
// Progress Reporting
// ---------------------------------------------------------------------------

/**
 * Phases of the analysis pipeline, used for progress reporting in the UI.
 */
export type AnalysisPhase =
  | "reading"
  | "connecting"
  | "analyzing"
  | "applying"
  | "done";

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
 * Type of tracked change operation, used by the Strategy pattern in wordApi.
 *
 * All types are applied via OOXML markup (`<w:del>`, `<w:ins>`) with an
 * attached Word comment containing the justification. This produces a rich
 * Review pane UI showing author, description, and reason for each change.
 *
 * - `"insert"` — Text insertion only (`<w:ins>` markup).
 * - `"delete"` — Text deletion only (`<w:del>` markup).
 * - `"replace"` — Combined deletion + insertion (`<w:del>` + `<w:ins>`).
 */
export type ChangeType = "insert" | "delete" | "replace";

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
