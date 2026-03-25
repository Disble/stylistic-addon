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
  ChunkSubmitResult,
  ChunkPollResult,
  ChunkAnalysisStatus,
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
   * Starts an asynchronous workflow run for a chunk and returns its `runId`.
   *
   * Note: retry logic is NOT here — it lives in `RetryAnalysisDecorator`.
   */
  async submitChunkAnalysis(
    chunk: TextChunk,
    genero: string,
    autorSlug: string
  ): Promise<ChunkSubmitResult> {
    const inputData: WorkflowInput = {
      text: chunk.text,
      genero: genero as WorkflowInput["genero"],
      autorSlug,
    };
    console.log(
      `🤖 [MastraAdapter] submitChunkAnalysis #${chunk.index} — ${chunk.text.length} chars, genero: "${genero}", autor: "${autorSlug}"`
    );

    try {
      const workflow = client.getWorkflow(WORKFLOW_ID);
      const run = await workflow.createRun();
      const runId = this.extractRunId(run);

      if (!runId) {
        return {
          chunkIndex: chunk.index,
          error: "Workflow createRun did not return a valid runId",
        };
      }

      await run.start({ inputData });

      console.log(`🚀 [MastraAdapter] Chunk #${chunk.index} enviado con runId "${runId}"`);
      return { chunkIndex: chunk.index, runId };
    } catch (error: unknown) {
      const message = this.normalizeErrorMessage(error);

      console.error(`💥 [MastraAdapter] Submit chunk #${chunk.index} error:`, message);
      return {
        chunkIndex: chunk.index,
        error: message,
      };
    }
  }

  /**
   * Polls a workflow run by `runId` and maps the workflow state into the
   * domain-level chunk polling contract.
   */
  async pollChunkAnalysis(chunkIndex: number, runId: string): Promise<ChunkPollResult> {
    console.log(`🔄 [MastraAdapter] Polling chunk #${chunkIndex} runId "${runId}"`);

    const workflow = client.getWorkflow(WORKFLOW_ID);
    const state = await workflow.runById(runId, {
      fields: ["result", "error"],
      withNestedWorkflows: false,
    });

    const normalizedState = this.validatePollState(state);

    if (!normalizedState) {
      return {
        chunkIndex,
        runId,
        status: "failed",
        suggestions: [],
        error: "Invalid workflow poll payload: missing status",
      };
    }

    console.log(`🔄 [MastraAdapter] Chunk #${chunkIndex} polled status: "${normalizedState.status}"`);

    if (normalizedState.status === "success") {
      const output = this.validateSuccessOutput(normalizedState.result);

      if (!output) {
        return {
          chunkIndex,
          runId,
          status: "failed",
          suggestions: [],
          error: "Invalid workflow success payload: expected suggestions[]",
        };
      }

      const suggestions = this.mapSuggestions(output.suggestions, chunkIndex);
      console.log(`✅ [MastraAdapter] Chunk #${chunkIndex} → ${suggestions.length} sugerencias`);
      return {
        chunkIndex,
        runId,
        status: "success",
        suggestions,
      };
    }

    if (this.requiresResume(normalizedState.status)) {
      return {
        chunkIndex,
        runId,
        status: "failed",
        suggestions: [],
        error: `Workflow entered \"${normalizedState.status}\" state and requires resume(), which this frontend does not support`,
      };
    }

    if (this.isNonTerminalStatus(normalizedState.status)) {
      return {
        chunkIndex,
        runId,
        status: normalizedState.status,
        suggestions: [],
      };
    }

    if (!this.isTerminalStatus(normalizedState.status)) {
      return {
        chunkIndex,
        runId,
        status: "failed",
        suggestions: [],
        error: `Unknown workflow status: ${normalizedState.status}`,
      };
    }

    return {
      chunkIndex,
      runId,
      status: normalizedState.status,
      suggestions: [],
      error: this.extractWorkflowError(normalizedState.error, normalizedState.status),
    };
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
    if (!this.isRecord(result)) {
      return undefined;
    }

    if (!Array.isArray(result.suggestions)) {
      return undefined;
    }

    if (!result.suggestions.every((suggestion) => this.isWorkflowSuggestion(suggestion))) {
      return undefined;
    }

    if (result.warnings !== undefined) {
      if (!Array.isArray(result.warnings) || !result.warnings.every((warning) => typeof warning === "string")) {
        return undefined;
      }
    }

    return {
      suggestions: result.suggestions,
      warnings: result.warnings,
    } as WorkflowOutput;
  }

  private isNonTerminalStatus(status: string): status is ChunkAnalysisStatus {
    return ["running", "pending", "waiting"].includes(status);
  }

  private requiresResume(status: string): boolean {
    return ["suspended", "paused"].includes(status);
  }

  private isTerminalStatus(status: string): status is Exclude<ChunkAnalysisStatus, "running" | "pending" | "waiting"> {
    return ["success", "failed", "tripwire", "canceled", "bailed"].includes(status);
  }

  private extractWorkflowError(error: unknown, status: string): string {
    const normalizedMessage = this.readErrorMessage(error);

    if (normalizedMessage) {
      return normalizedMessage;
    }

    if (status === "success") {
      return "Workflow reported success without a valid result payload";
    }

    return `Workflow terminated with status "${status}" without an error payload`;
  }

  private normalizeErrorMessage(error: unknown): string {
    const normalizedMessage = this.readErrorMessage(error);

    if (normalizedMessage) {
      return normalizedMessage;
    }

    return "Unknown analysis error";
  }

  private validatePollState(state: unknown): { status: string; result?: unknown; error?: unknown } | undefined {
    if (!this.isRecord(state)) {
      return undefined;
    }

    const status = this.readNonEmptyString(state.status);

    if (!status) {
      return undefined;
    }

    return {
      status,
      result: state.result,
      error: state.error,
    };
  }

  private extractRunId(run: unknown): string | undefined {
    if (!this.isRecord(run)) {
      return undefined;
    }

    return this.readNonEmptyString(run.runId);
  }

  private isWorkflowSuggestion(value: unknown): value is WorkflowSuggestion {
    if (!this.isRecord(value)) {
      return false;
    }

    return (
      this.readNonEmptyString(value.originalText) !== undefined &&
      this.readNonEmptyString(value.suggestedText) !== undefined &&
      this.readNonEmptyString(value.justification) !== undefined &&
      this.readNonEmptyString(value.category) !== undefined &&
      this.isSuggestionSeverity(value.severity)
    );
  }

  private isSuggestionSeverity(value: unknown): value is WorkflowSuggestion["severity"] {
    return value === "high" || value === "medium" || value === "low";
  }

  private readErrorMessage(value: unknown): string | undefined {
    if (value instanceof Error) {
      return this.readNonEmptyString(value.message);
    }

    if (typeof value === "string") {
      return this.readNonEmptyString(value);
    }

    if (!this.isRecord(value)) {
      return undefined;
    }

    const directMessage = this.readNonEmptyString(value.message);

    if (directMessage) {
      return directMessage;
    }

    try {
      return JSON.stringify(value);
    } catch {
      return undefined;
    }
  }

  private readNonEmptyString(value: unknown): string | undefined {
    if (typeof value !== "string") {
      return undefined;
    }

    const normalized = value.trim();
    return normalized.length > 0 ? normalized : undefined;
  }

  private isRecord(value: unknown): value is Record<string, unknown> {
    return typeof value === "object" && value !== null;
  }
}
