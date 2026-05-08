/**
 * Pipeline execution context — the shared state that flows through all
 * Chain of Responsibility handlers.
 *
 * Each handler reads from and writes to this context. The `aborted` flag
 * allows any handler to short-circuit the chain without throwing.
 *
 * Immutable fields (set at construction) are separated from mutable fields
 * (set by handlers) for clarity.
 *
 * @module PipelineContext
 */

import type { TextChunk } from "../chunking/TextChunk.types";
import type { ApplySuggestionsResult } from "../DocumentApplication.types";
import type { AnalysisProfileId } from "../Profile.types";
import type { IAnalysisPort, IDocumentPort } from "../ports";
import type { Suggestion } from "../suggestion/Suggestion.types";
import type { PipelineEventEmitter } from "./PipelineEvents";

/**
 * The full context object passed between pipeline handlers.
 *
 * Immutable fields (set at pipeline start, never changed by handlers):
 * - `documentPort`, `analysisPort` — the injected adapters
 * - `emitter` — the event emitter for progress/lifecycle notifications
 * - `genero` — the selected genre for the stylistic workflow
 * - `maxChunkSize` — character limit per chunk
 *
 * Mutable fields (populated progressively by handlers):
 * - `text`, `isSelection`, `documentUuid` — set by ReadTextHandler
 * - `chunks` — set by ChunkTextHandler
 * - `rawSuggestions`, `chunkErrors` — set by AnalyzeChunksHandler
 * - `uniqueSuggestions` — set by DeduplicateHandler
 * - `pendingSuggestions` — set by GuardAppliedHandler
 * - `result` — set by ApplySuggestionsHandler
 * - `aborted`, `abortReason` — set by any handler to stop the chain
 */
export interface PipelineContext {
  // --- Injected at pipeline start (immutable) ---

  /** Port for all document operations (read text, apply suggestions, cleanup). */
  readonly documentPort: IDocumentPort;

  /** Port for all analysis backend operations (check connection, analyze chunks). */
  readonly analysisPort: IAnalysisPort;

  /** Event emitter for progress notifications (Observer pattern). */
  readonly emitter: PipelineEventEmitter;

  /** The genre selected by the user, sent to the stylistic workflow (e.g., "narrativa-literaria", "general"). */
  readonly genero: AnalysisProfileId;

  /** Maximum characters per chunk sent to the backend. */
  readonly maxChunkSize: number;

  // --- Set by ReadTextHandler ---
  text?: string;
  isSelection?: boolean;
  documentUuid?: string;

  // --- Set by ChunkTextHandler ---
  chunks?: TextChunk[];

  // --- Set by AnalyzeChunksHandler ---
  rawSuggestions?: Suggestion[];
  chunkErrors?: string[];

  // --- Set by DeduplicateHandler ---
  uniqueSuggestions?: Suggestion[];

  // --- Set by GuardAppliedHandler ---
  pendingSuggestions?: Suggestion[];

  // --- Set by ApplySuggestionsHandler ---
  result?: ApplySuggestionsResult;

  // --- Abort mechanism (set by any handler) ---

  /**
   * When `true`, subsequent handlers in the chain will be skipped.
   * Set by any handler that wants to abort the pipeline gracefully
   * (e.g., empty document, backend unavailable, no suggestions).
   */
  aborted?: boolean;

  /**
   * Human-readable reason for abortion, shown in the UI status bar.
   * Always set together with `aborted = true`.
   */
  abortReason?: string;
}
