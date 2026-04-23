/**
 * Port interfaces (Hexagonal Architecture) — define the contracts between
 * the domain/application layer and the infrastructure adapters.
 *
 * These interfaces are the only cross-boundary contracts in the system:
 * - `IDocumentPort` — abstracts all Word document operations.
 * - `IAnalysisPort` — abstracts all Mastra backend operations.
 *
 * Benefits:
 * - The pipeline (`domain/pipeline/`) depends only on these interfaces.
 * - Adapters (`adapters/`) implement these interfaces.
 * - Tests can inject mock implementations without touching Office.js or Mastra.
 * - Swapping the backend (e.g., OpenAI direct) or document host (e.g., Google Docs)
 *   requires only a new adapter — zero changes to the pipeline or UI.
 *
 * @module ports
 */

import type {
  ApplySuggestionsResult,
  ChunkPollResult,
  ChunkSubmitResult,
  DocumentReviewState,
  FeedbackPayload,
  ProgressCallback,
  ResolutionTelemetryEvent,
  Suggestion,
  SuggestionActionResult,
  TextChunk,
  TextSource,
} from "./types";

// ---------------------------------------------------------------------------
// Document Port
// ---------------------------------------------------------------------------

/**
 * Contract for all document read/write operations.
 *
 * Implemented by `WordAdapter` for Microsoft Word via Office.js.
 * Could be implemented for Google Docs, LibreOffice, etc. without touching
 * any pipeline or UI code.
 */
export interface IDocumentPort {
  /**
   * Resolves the text to analyze: returns the current selection if non-empty,
   * otherwise falls back to the full document body.
   */
  getTextToAnalyze(): Promise<TextSource>;

  /**
   * Returns the set of original texts already applied as Stylistic tracked
   * deletions in the document. Used as a guard to prevent duplicate tracked
   * changes on re-run.
   */
  getAppliedOriginalTexts(): Promise<Set<string>>;

  /**
   * Applies an array of suggestions as native tracked changes with embedded
   * justification comments, one suggestion at a time.
   *
   * @param suggestions - Suggestions to apply.
   * @param onProgress  - Optional progress callback invoked after each suggestion.
   */
  applySuggestions(
    suggestions: Suggestion[],
    onProgress?: ProgressCallback,
  ): Promise<ApplySuggestionsResult>;

  /**
   * Returns a dry-run summary of the cleanup operation.
   * Used by the task pane to decide whether the cleanup CTA should be visible.
   */
  getCleanupPreview(): Promise<{ deletable: number; kept: number }>;

  /**
   * Deletes Stylistic comments whose tracked changes have been resolved.
   * Never touches comments from other authors.
   */
  cleanupResolvedComments(): Promise<{ deleted: number; kept: number }>;

  /**
   * Accepts all Stylistic tracked changes associated with a suggestion.
   * Also deletes the associated Stylistic comment if present.
   * Returns a result object — never throws.
   *
   * Result contract:
   * - `accepted` means tracked changes and Stylistic cleanup completed atomically
   * - `unobservable` means Word could not prove the suggestion state yet
   * - `identity-lost` means compound v2 metadata exists but is incomplete/corrupt
   */
  acceptSuggestion(suggestion: Suggestion): Promise<SuggestionActionResult>;

  /**
   * Rejects all Stylistic tracked changes associated with a suggestion.
   * Also deletes the associated Stylistic comment if present.
   * Returns a result object — never throws.
   *
   * Result contract:
   * - `rejected` means tracked changes and Stylistic cleanup completed atomically
   * - `unobservable` means Word could not prove the suggestion state yet
   * - `identity-lost` means compound v2 metadata exists but is incomplete/corrupt
   */
  rejectSuggestion(suggestion: Suggestion): Promise<SuggestionActionResult>;

  /**
   * Inspects the document-derived Stylistic review state.
   * The document is the source of truth for pending artifacts.
   */
  getDocumentReviewState(): Promise<DocumentReviewState>;

  /**
   * Lets the user explicitly disable Word Track Changes after Stylistic pending
   * artifacts reach zero. Never called automatically.
   */
  disableTrackChanges(): Promise<void>;

  /**
   * Navigates the document view to the real Word artifact for one suggestion.
   * Prefers persisted Stylistic identity and falls back to text search only when
   * the artifact can no longer be re-located directly.
   * Never throws — silently no-ops if the target is not found.
   */
  navigateToText(target: Suggestion | string): Promise<void>;
}

// ---------------------------------------------------------------------------
// Analysis Port
// ---------------------------------------------------------------------------

/**
 * Contract for all AI analysis backend operations.
 *
 * Implemented by `MastraAdapter` (wrapped with `RetryAnalysisDecorator`).
 * Could be implemented for any LLM backend or local analysis engine.
 */
export interface IAnalysisPort {
  /**
   * Checks whether the backend is reachable and the workflow exists.
   * Used as a fail-fast gate before starting analysis.
   * Never throws — returns `false` on any error.
   */
  checkConnection(): Promise<boolean>;

  /**
   * Submits a text chunk for asynchronous workflow execution.
   * Returns the `runId` required for later polling.
   */
  submitChunkAnalysis(
    chunk: TextChunk,
    genero: string,
    autorSlug: string,
  ): Promise<ChunkSubmitResult>;

  /**
   * Polls an existing workflow run created by `submitChunkAnalysis()`.
   * Returns intermediate or terminal workflow state for the chunk.
   */
  pollChunkAnalysis(
    chunkIndex: number,
    runId: string,
  ): Promise<ChunkPollResult>;
}

// ---------------------------------------------------------------------------
// Feedback Port
// ---------------------------------------------------------------------------

/**
 * Contract for sending user feedback about suggestions.
 *
 * Implemented by `FeedbackAdapter`.
 * Fire-and-forget — never awaited in the UI. Errors must be swallowed silently.
 */
export interface IFeedbackPort {
  /**
   * Sends a feedback payload to the backend.
   * Must execute asynchronously; errors must be swallowed silently.
   */
  sendFeedback(payload: FeedbackPayload): Promise<void>;
}

// ---------------------------------------------------------------------------
// Telemetry Port
// ---------------------------------------------------------------------------

/**
 * Contract for best-effort observability events emitted by workflow phases.
 *
 * Telemetry must NEVER alter semantic document outcomes. Adapters are expected
 * to swallow transport/storage failures internally or let callers degrade them
 * to warnings.
 */
export interface ITelemetryPort {
  /** Emits one structured telemetry event for later debugging or analysis. */
  emit(event: ResolutionTelemetryEvent): Promise<void>;
}
