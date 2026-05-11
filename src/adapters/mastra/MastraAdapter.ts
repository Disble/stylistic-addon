/* global console */

/**
 * Mastra Adapter — implements `IAnalysisPort` using `@mastra/client-js`.
 *
 * Wraps the Mastra workflow client for the editorial analysis backend.
 * Focused exclusively on communication: no retry logic (handled by
 * `RetryAnalysisDecorator`), no DOM interaction, no Office.js.
 *
 * Mastra clients are created through `MastraClientFactory` so every request can
 * include the latest Better Auth bearer token.
 *
 * @module MastraAdapter
 */

import type { TextChunk } from "../../domain/chunking/TextChunk.types";
import type {
  ChunkCancelResult,
  ChunkAnalysisStatus,
  ChunkPollResult,
  ChunkRunReference,
  ChunkSubmitResult,
  WorkflowInput,
  WorkflowOutput,
  WorkflowSubmitContext,
  WorkflowSuggestion,
} from "../../domain/mastra/MastraWorkflow.types";
import type { IAnalysisPort } from "../../domain/ports";
import type { Suggestion } from "../../domain/suggestion/Suggestion.types";
import {
  MASTRA_BASE_URL,
  MASTRA_POLL_BYPASS_ENABLED,
  WORKFLOW_ID,
} from "../../infrastructure/config";
import { MastraClientFactory } from "./MastraClientFactory";
import { createMockMastraPollOutput } from "./MockMastraPollOutputFactory";

/** Bridges the analysis port to the Mastra workflow client. */
export class MastraAdapter implements IAnalysisPort {
  constructor(private readonly clientFactory = new MastraClientFactory()) {}

  /**
   * Checks whether the Mastra backend is reachable and the editorial
   * workflow is registered. Never throws — returns `false` on any error.
   */
  async checkConnection(): Promise<boolean> {
    if (MASTRA_POLL_BYPASS_ENABLED) {
      console.log(
        "🧪 [MastraAdapter] Bypass activo: se omite checkConnection real y se asume conexión OK"
      );
      return true;
    }

    try {
      console.log(
        `🔌 [MastraAdapter] Verificando → ${MASTRA_BASE_URL}, workflow: "${WORKFLOW_ID}"`
      );
      const workflow = this.clientFactory.create().getWorkflow(WORKFLOW_ID);
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
    input: WorkflowSubmitContext
  ): Promise<ChunkSubmitResult> {
    if (MASTRA_POLL_BYPASS_ENABLED) {
      return this.buildBypassedSubmitResult(chunk.index);
    }

    const inputData: WorkflowInput = {
      text: chunk.text,
      ...input,
    };

    const generoLabel = input.genero ?? "sin-genero";
    console.log(
      `🤖 [MastraAdapter] submitChunkAnalysis #${chunk.index} — ${chunk.text.length} chars, genero: "${generoLabel}", documentUuid: "${input.documentUuid}"`
    );

    try {
      const workflow = this.clientFactory.create().getWorkflow(WORKFLOW_ID);
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

    if (MASTRA_POLL_BYPASS_ENABLED) {
      return this.buildBypassedPollResult(chunkIndex, runId);
    }

    const workflow = this.clientFactory.create().getWorkflow(WORKFLOW_ID);
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
        origin: "frontend-terminal",
        suggestions: [],
        error: "Invalid workflow poll payload: missing status",
      };
    }

    console.log(
      `🔄 [MastraAdapter] Chunk #${chunkIndex} polled status: "${normalizedState.status}"`
    );

    if (normalizedState.status === "success") {
      const output = this.validateSuccessOutput(normalizedState.result);

      if (!output) {
        return {
          chunkIndex,
          runId,
          status: "failed",
          origin: "frontend-terminal",
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
        origin: "backend",
        suggestions,
      };
    }

    if (this.requiresResume(normalizedState.status)) {
      return {
        chunkIndex,
        runId,
        status: "failed",
        origin: "backend",
        suggestions: [],
        error: `Workflow entered "${normalizedState.status}" state and requires resume(), which this frontend does not support`,
      };
    }

    if (this.isNonTerminalStatus(normalizedState.status)) {
      return {
        chunkIndex,
        runId,
        status: normalizedState.status,
        origin: "backend",
        suggestions: [],
      };
    }

    if (!this.isTerminalStatus(normalizedState.status)) {
      return {
        chunkIndex,
        runId,
        status: "failed",
        origin: "frontend-terminal",
        suggestions: [],
        error: `Unknown workflow status: ${normalizedState.status}`,
      };
    }

    return {
      chunkIndex,
      runId,
      status: normalizedState.status,
      origin: "backend",
      suggestions: [],
      error: this.extractWorkflowError(normalizedState.error, normalizedState.status),
    };
  }

  /** Cancels one existing Mastra run by rehydrating it from the known `runId`. */
  async cancelChunkAnalysis(chunkIndex: number, runId: string): Promise<ChunkCancelResult> {
    if (MASTRA_POLL_BYPASS_ENABLED) {
      return { chunkIndex, runId, canceled: true };
    }

    try {
      const workflow = this.clientFactory.create().getWorkflow(WORKFLOW_ID);
      const run = await workflow.createRun({ runId });
      await run.cancel();
      return { chunkIndex, runId, canceled: true };
    } catch (error: unknown) {
      return {
        chunkIndex,
        runId,
        canceled: false,
        error: this.normalizeErrorMessage(error),
      };
    }
  }

  /** Re-polls the same backend run without creating a new submission. */
  retryPollChunkAnalysis(reference: ChunkRunReference): Promise<ChunkPollResult> {
    return this.pollChunkAnalysis(reference.chunkIndex, reference.runId);
  }

  /** Maps raw workflow suggestions to `Suggestion` objects with assigned IDs. */
  private mapSuggestions(raw: WorkflowSuggestion[] | undefined, chunkIndex: number): Suggestion[] {
    if (!raw || !Array.isArray(raw)) return [];
    return raw.map((s, i) => {
      const type = s.type ?? "track-change";
      const suggestion: Suggestion = {
        id: `chunk${chunkIndex}-${i}`,
        context: s.context,
        anchor: s.anchor,
        justification: s.justification,
        category: s.category,
        severity: s.severity,
        type,
      };
      if (type !== "comment-only") {
        suggestion.suggestedText = s.suggestedText;
      }
      return suggestion;
    });
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
      if (
        !Array.isArray(result.warnings) ||
        !result.warnings.every((warning) => typeof warning === "string")
      ) {
        return undefined;
      }
    }

    return {
      suggestions: result.suggestions,
      warnings: result.warnings,
    } as WorkflowOutput;
  }

  /** Builds a deterministic success result for development poll bypass mode. */
  private buildBypassedPollResult(chunkIndex: number, runId: string): ChunkPollResult {
    // The bypass fixture represents one backend response, not one response per
    // chunk. Returning it only for chunk 0 keeps multi-chunk document tests from
    // multiplying the same mock suggestions into hundreds of Word mutations.
    if (chunkIndex !== 0) {
      console.log(
        `🧪 [MastraAdapter] Poll bypass activo para chunk #${chunkIndex} → 0 sugerencias mockeadas`
      );

      return {
        chunkIndex,
        runId,
        status: "success",
        origin: "backend",
        suggestions: [],
      };
    }

    const suggestions = this.mapSuggestions(createMockMastraPollOutput().suggestions, chunkIndex);
    console.log(
      `🧪 [MastraAdapter] Poll bypass activo para chunk #${chunkIndex} → ${suggestions.length} sugerencias mockeadas`
    );

    return {
      chunkIndex,
      runId,
      status: "success",
      origin: "backend",
      suggestions,
    };
  }

  /** Builds a deterministic submit result for development bypass mode. */
  private buildBypassedSubmitResult(chunkIndex: number): ChunkSubmitResult {
    const runId = `bypass-run-${chunkIndex}`;
    console.log(
      `🧪 [MastraAdapter] Submit bypass activo para chunk #${chunkIndex} → runId "${runId}"`
    );

    return {
      chunkIndex,
      runId,
    };
  }

  private isNonTerminalStatus(status: string): status is ChunkAnalysisStatus {
    return ["running", "pending", "waiting"].includes(status);
  }

  private requiresResume(status: string): boolean {
    return ["suspended", "paused"].includes(status);
  }

  private isTerminalStatus(
    status: string
  ): status is Exclude<ChunkAnalysisStatus, "running" | "pending" | "waiting"> {
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

  private validatePollState(
    state: unknown
  ): { status: string; result?: unknown; error?: unknown } | undefined {
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

    const context = this.readNonEmptyString(value.context);
    const anchor = this.readNonEmptyString(value.anchor);

    if (
      context === undefined ||
      anchor === undefined ||
      this.readNonEmptyString(value.justification) === undefined ||
      this.readNonEmptyString(value.category) === undefined ||
      !this.isSuggestionSeverity(value.severity)
    ) {
      return false;
    }

    if (!context.includes(anchor)) {
      return false;
    }

    // `type` is optional; when present it must be a known kind
    if (value.type !== undefined && !this.isSuggestionType(value.type)) {
      return false;
    }

    const type = value.type ?? "track-change";

    // `suggestedText` is required only for track-change suggestions. Empty
    // string is valid for delete-only native Track Changes; markdown strings are
    // also valid because apply-time formatting is decoded from transport text.
    if (type === "track-change" && typeof value.suggestedText !== "string") {
      return false;
    }

    return true;
  }

  private isSuggestionType(value: unknown): value is WorkflowSuggestion["type"] {
    return value === "track-change" || value === "comment-only";
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
