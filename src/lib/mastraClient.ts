/**
 * Mastra workflow client — wraps `@mastra/client-js` for communication
 * with the editorial analysis backend.
 *
 * Responsibilities:
 * - Initialize a singleton {@link MastraClient} instance.
 * - Execute the editorial workflow per text chunk with structured input.
 * - Retry transient failures with exponential backoff.
 * - Validate backend connectivity before starting analysis.
 *
 * This module never touches Office.js or the DOM (SRP). It depends only
 * on {@link types}, {@link config}, and `@mastra/client-js`.
 *
 * @module mastraClient
 */

import { MastraClient } from "@mastra/client-js";
import {
  MASTRA_BASE_URL,
  WORKFLOW_ID,
  MAX_RETRIES,
  RETRY_BASE_DELAY_MS,
} from "./config";
import {
  TextChunk,
  Suggestion,
  ChunkResult,
  WorkflowInput,
  WorkflowOutput,
  WorkflowSuggestion,
} from "./types";

// ---------------------------------------------------------------------------
// Singleton client
// ---------------------------------------------------------------------------

/** Singleton Mastra client instance, reused across all calls. */
const client = new MastraClient({ baseUrl: MASTRA_BASE_URL });

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Checks whether the Mastra backend is reachable and the editorial
 * workflow exists.
 *
 * Used as a fail-fast gate before starting a potentially long analysis.
 * Does not throw — returns `false` on any error.
 *
 * @returns `true` if the workflow is accessible, `false` otherwise.
 */
export async function checkConnection(): Promise<boolean> {
  try {
    const workflow = client.getWorkflow(WORKFLOW_ID);
    await workflow.details();
    return true;
  } catch {
    return false;
  }
}

/**
 * Sends a text chunk to the Mastra editorial workflow and returns
 * the resulting suggestions.
 *
 * On transient failures (network errors, timeouts, workflow "failed" status),
 * retries up to {@link MAX_RETRIES} times with exponential backoff.
 * On permanent failure, returns an empty suggestion list with an error message.
 *
 * **Never throws.** The caller always receives a {@link ChunkResult}.
 *
 * @param chunk   - The text chunk to analyze (with positional metadata).
 * @param profile - Analysis profile identifier (e.g., "general", "formal").
 * @returns A {@link ChunkResult} containing suggestions or an error.
 */
export async function analyzeChunk(
  chunk: TextChunk,
  profile: string
): Promise<ChunkResult> {
  const inputData: WorkflowInput = {
    text: chunk.text,
    profile,
    chunkIndex: chunk.index,
    totalChunks: chunk.total,
  };

  let lastError = "";

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    if (attempt > 0) {
      await delay(RETRY_BASE_DELAY_MS * Math.pow(2, attempt - 1));
    }

    try {
      const workflow = client.getWorkflow(WORKFLOW_ID);
      const run = await workflow.createRun();
      const result = await run.start({ inputData });

      if (result.status === "success") {
        const output = result.result as WorkflowOutput;
        const suggestions = mapSuggestions(output.suggestions, chunk.index);
        return { chunkIndex: chunk.index, suggestions };
      }

      // Workflow returned a non-success status — treat as retryable
      lastError = `Workflow status: ${result.status}`;
    } catch (error: unknown) {
      lastError = error instanceof Error ? error.message : String(error);
    }
  }

  // All retries exhausted
  return {
    chunkIndex: chunk.index,
    suggestions: [],
    error: `Chunk ${chunk.index + 1}: ${lastError}`,
  };
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Maps raw workflow suggestions to frontend {@link Suggestion} objects
 * by assigning unique IDs scoped to the chunk.
 *
 * @param raw        - Suggestions from the workflow response.
 * @param chunkIndex - Index of the chunk these suggestions belong to.
 * @returns Suggestions with assigned `id` fields.
 */
function mapSuggestions(
  raw: WorkflowSuggestion[] | undefined,
  chunkIndex: number
): Suggestion[] {
  if (!raw || !Array.isArray(raw)) {
    return [];
  }

  return raw.map((s, i) => ({
    id: `chunk${chunkIndex}-${i}`,
    originalText: s.originalText,
    suggestedText: s.suggestedText,
    justification: s.justification,
    category: s.category,
  }));
}

/**
 * Returns a promise that resolves after the specified duration.
 * Used for exponential backoff between retry attempts.
 *
 * @param ms - Delay in milliseconds.
 */
function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
