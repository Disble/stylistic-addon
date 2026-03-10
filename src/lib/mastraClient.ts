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
import { MASTRA_BASE_URL, WORKFLOW_ID, MAX_RETRIES, RETRY_BASE_DELAY_MS } from "./config";
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
    console.log(
      `🔌 [MastraClient] Verificando conexión → ${MASTRA_BASE_URL}, workflow: "${WORKFLOW_ID}"`
    );
    const workflow = client.getWorkflow(WORKFLOW_ID);
    await workflow.details();
    console.log("🔌 [MastraClient] ✅ Conexión exitosa");
    return true;
  } catch (err) {
    console.error("🔌 [MastraClient] ❌ Conexión fallida:", err);
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
  profile: string,
  language: string
): Promise<ChunkResult> {
  const inputData: WorkflowInput = {
    text: chunk.text,
    profile,
    language,
  };

  console.log(
    `🤖 [MastraClient] analyzeChunk #${chunk.index} — ${chunk.text.length} chars, perfil: "${profile}", idioma: "${language}"`
  );
  let lastError = "";

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    if (attempt > 0) {
      const delayMs = RETRY_BASE_DELAY_MS * Math.pow(2, attempt - 1);
      console.warn(
        `🔄 [MastraClient] Retry ${attempt}/${MAX_RETRIES} para chunk #${chunk.index} — esperando ${delayMs}ms`
      );
      await delay(delayMs);
    }

    try {
      console.log(`🤖 [MastraClient] Chunk #${chunk.index} intento ${attempt}: creando run...`);
      const workflow = client.getWorkflow(WORKFLOW_ID);
      const run = await workflow.createRun();
      console.log(
        `🤖 [MastraClient] Chunk #${chunk.index} intento ${attempt}: ejecutando workflow...`
      );
      const result = await run.startAsync({ inputData });
      console.log(
        `🤖 [MastraClient] Chunk #${chunk.index} intento ${attempt}: status = "${result.status}"`
      );

      if (result.status === "success") {
        const output = result.result as WorkflowOutput;
        console.log(
          `✅ [MastraClient] Chunk #${chunk.index} → ${output.suggestions?.length ?? 0} sugerencias raw del workflow`
        );
        const suggestions = mapSuggestions(output.suggestions, chunk.index);
        return { chunkIndex: chunk.index, suggestions };
      }

      // Workflow returned a non-success status — treat as retryable
      console.warn(`⚠️ [MastraClient] Chunk #${chunk.index} status no exitoso: "${result.status}"`);
      lastError = `Workflow status: ${result.status}`;
    } catch (error: unknown) {
      lastError = error instanceof Error ? error.message : String(error);
      console.error(`💥 [MastraClient] Chunk #${chunk.index} intento ${attempt} error:`, lastError);
    }
  }

  // All retries exhausted
  console.error(
    `❌ [MastraClient] Chunk #${chunk.index} agotó ${MAX_RETRIES} reintentos. Último error: ${lastError}`
  );
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
function mapSuggestions(raw: WorkflowSuggestion[] | undefined, chunkIndex: number): Suggestion[] {
  if (!raw || !Array.isArray(raw)) {
    console.warn(`⚠️ [MastraClient] mapSuggestions: chunk #${chunkIndex} sin sugerencias raw`);
    return [];
  }
  console.log(`🗺️ [MastraClient] Mapeando ${raw.length} sugerencias para chunk #${chunkIndex}`);

  return raw.map((s, i) => ({
    id: `chunk${chunkIndex}-${i}`,
    originalText: s.originalText,
    suggestedText: s.suggestedText,
    justification: s.justification,
    category: s.category,
    severity: s.severity,
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
