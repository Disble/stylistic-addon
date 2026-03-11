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

import {
  TextSource,
  Suggestion,
  InsertionResult,
  ProgressCallback,
  TextChunk,
  ChunkResult,
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
    onProgress?: ProgressCallback
  ): Promise<InsertionResult>;

  /**
   * Deletes Stylistic comments whose tracked changes have been resolved.
   * Never touches comments from other authors.
   */
  cleanupResolvedComments(): Promise<{ deleted: number; kept: number }>;
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
   * Sends a text chunk to the analysis backend and returns suggestions.
   * Never throws — returns `ChunkResult` with empty suggestions on failure.
   *
   * @param chunk    - The text chunk with positional metadata.
   * @param profile  - Analysis profile (e.g., "general", "formal").
   * @param language - ISO 639-1 language code (e.g., "es").
   */
  analyzeChunk(chunk: TextChunk, profile: string, language: string): Promise<ChunkResult>;
}
