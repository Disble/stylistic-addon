/* global console */

/**
 * Mastra Adapter — implements `IAnalysisPort` using `@mastra/client-js`.
 *
 * Wraps the Mastra workflow client for the editorial analysis backend.
 * Focused exclusively on communication: no retry logic (handled by
 * `RetryAnalysisDecorator`), no DOM interaction, no Office.js.
 *
 * The singleton `MastraClient` instance is reused across all calls for
 * connection efficiency.
 *
 * @module MastraAdapter
 */

import { MastraClient } from "@mastra/client-js";
import { IAnalysisPort } from "../../domain/ports";
import {
  TextChunk,
  ChunkResult,
  Suggestion,
  WorkflowInput,
  WorkflowOutput,
  WorkflowSuggestion,
} from "../../domain/types";
import { MASTRA_BASE_URL, WORKFLOW_ID } from "../../infrastructure/config";

/** Singleton Mastra client instance, reused across all calls. */
const client = new MastraClient({ baseUrl: MASTRA_BASE_URL });

export class MastraAdapter implements IAnalysisPort {
  /**
   * Checks whether the Mastra backend is reachable and the editorial
   * workflow is registered. Never throws — returns `false` on any error.
   */
  async checkConnection(): Promise<boolean> {
    try {
      console.log(
        `🔌 [MastraAdapter] Verificando → ${MASTRA_BASE_URL}, workflow: "${WORKFLOW_ID}"`
      );
      const workflow = client.getWorkflow(WORKFLOW_ID);
      await workflow.details();
      console.log("🔌 [MastraAdapter] ✅ Conexión exitosa");
      return true;
    } catch (err) {
      console.error("🔌 [MastraAdapter] ❌ Conexión fallida:", err);
      return false;
    }
  }

  /**
   * Sends a text chunk to the Mastra editorial workflow and returns
   * the resulting suggestions. Never throws — returns an empty
   * `ChunkResult` with an error message on failure.
   *
   * Note: retry logic is NOT here — it lives in `RetryAnalysisDecorator`.
   */
  async analyzeChunk(chunk: TextChunk, profile: string, language: string): Promise<ChunkResult> {
    const inputData: WorkflowInput = { text: chunk.text, profile, language };
    console.log(
      `🤖 [MastraAdapter] analyzeChunk #${chunk.index} — ${chunk.text.length} chars, perfil: "${profile}"`
    );

    try {
      const workflow = client.getWorkflow(WORKFLOW_ID);
      const run = await workflow.createRun();
      const result = await run.startAsync({ inputData });
      console.log(`🤖 [MastraAdapter] Chunk #${chunk.index} status: "${result.status}"`);

      if (result.status === "success") {
        const output = this.validateSuccessOutput(result.result);

        if (!output) {
          return {
            chunkIndex: chunk.index,
            suggestions: [],
            error: "Invalid workflow success payload",
          };
        }

        const suggestions = this.mapSuggestions(output.suggestions, chunk.index);
        console.log(`✅ [MastraAdapter] Chunk #${chunk.index} → ${suggestions.length} sugerencias`);
        return { chunkIndex: chunk.index, suggestions };
      }

      return {
        chunkIndex: chunk.index,
        suggestions: [],
        error: `Workflow status: ${result.status}`,
      };
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      console.error(`💥 [MastraAdapter] Chunk #${chunk.index} error:`, message);
      return {
        chunkIndex: chunk.index,
        suggestions: [],
        error: message,
      };
    }
  }

  /** Maps raw workflow suggestions to `Suggestion` objects with assigned IDs. */
  private mapSuggestions(raw: WorkflowSuggestion[] | undefined, chunkIndex: number): Suggestion[] {
    if (!raw || !Array.isArray(raw)) return [];
    return raw.map((s, i) => ({
      id: `chunk${chunkIndex}-${i}`,
      originalText: s.originalText,
      suggestedText: s.suggestedText,
      justification: s.justification,
      category: s.category,
      severity: s.severity,
    }));
  }

  private validateSuccessOutput(result: unknown): WorkflowOutput | undefined {
    if (!result || typeof result !== "object") {
      return undefined;
    }

    if (!("suggestions" in result)) {
      return undefined;
    }

    return result as WorkflowOutput;
  }
}
