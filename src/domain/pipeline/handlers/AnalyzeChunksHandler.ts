/* global console, setTimeout */

/**
 * AnalyzeChunksHandler — Phase 4 of the analysis pipeline.
 *
 * Sends each text chunk to the analysis backend sequentially and collects
 * suggestions. Applies the Partial Success philosophy: if some chunks fail,
 * suggestions from successful chunks are still applied.
 *
 * Progress is emitted per chunk via the PipelineEventEmitter.
 *
 * @module AnalyzeChunksHandler
 */

import { DEFAULT_AUTHOR_SLUG, POLL_INTERVAL_MS } from "../../../infrastructure/config";
import type { Suggestion } from "../../suggestion/Suggestion.types";
import type { PipelineContext } from "../PipelineContext";
import type { PipelineHandler } from "./ReadTextHandler";

type AnalyzeChunk = NonNullable<PipelineContext["chunks"]>[number];

type AnalysisState = {
  scope: string;
  chunks: NonNullable<PipelineContext["chunks"]>;
  allSuggestions: Suggestion[];
  chunkErrors: string[];
  pendingRuns: Map<number, string>;
  totalSteps: number;
  submittedCount: number;
  completedCount: number;
};

export class AnalyzeChunksHandler implements PipelineHandler {
  constructor(private readonly pollIntervalMs: number = POLL_INTERVAL_MS) {}

  /** Coordinates chunk submission, polling, and abort handling for analysis. */
  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    const state = this.createAnalysisState(ctx);

    console.log(
      `🤖 [AnalyzeChunksHandler] Fase 4: Analizando ${state.chunks.length} chunk(s) con Mastra...`
    );

    await this.submitChunks(ctx, state);
    await this.collectPendingRuns(ctx, state);
    this.assignResults(ctx, state);

    if (this.shouldAbort(state)) {
      this.abortAnalysis(ctx, state.chunkErrors.length);
      return;
    }

    await next();
  }

  /** Builds the mutable state shared by submission and polling phases. */
  private createAnalysisState(ctx: PipelineContext): AnalysisState {
    const chunks = ctx.chunks!;

    return {
      scope: ctx.isSelection ? "selección" : "documento",
      chunks,
      allSuggestions: [],
      chunkErrors: [],
      pendingRuns: new Map<number, string>(),
      totalSteps: Math.max(chunks.length * 2, 1),
      submittedCount: 0,
      completedCount: 0,
    };
  }

  /** Submits every chunk to the backend and records pending runs or submit errors. */
  private async submitChunks(ctx: PipelineContext, state: AnalysisState): Promise<void> {
    for (const chunk of state.chunks) {
      await this.submitChunk(ctx, state, chunk);
    }
  }

  /** Submits a single chunk and updates progress bookkeeping for the queue phase. */
  private async submitChunk(
    ctx: PipelineContext,
    state: AnalysisState,
    chunk: AnalyzeChunk
  ): Promise<void> {
    ctx.emitter.emitPhaseStart(
      "analyzing",
      `Encolando fragmento ${chunk.index + 1} de ${state.chunks.length} (${state.scope})...`
    );
    ctx.emitter.emitProgress(
      state.submittedCount,
      state.totalSteps,
      `Encolando fragmento ${chunk.index + 1} de ${state.chunks.length}...`
    );

    console.log(
      `🤖 [AnalyzeChunksHandler] Encolando chunk ${chunk.index + 1}/${state.chunks.length} (${chunk.text.length} chars)`
    );

    const submitResult = await ctx.analysisPort.submitChunkAnalysis(
      chunk,
      ctx.genero,
      DEFAULT_AUTHOR_SLUG
    );
    state.submittedCount += 1;

    console.log(this.getSubmitLogMessage(chunk.index, submitResult.runId, submitResult.error));

    ctx.emitter.emitProgress(
      state.submittedCount,
      state.totalSteps,
      this.getSubmissionProgressMessage(chunk.index, Boolean(submitResult.runId))
    );

    if (submitResult.runId) {
      state.pendingRuns.set(chunk.index, submitResult.runId);
      return;
    }

    if (submitResult.error) {
      state.chunkErrors.push(submitResult.error);
    }
  }

  /** Polls all pending runs until every chunk reaches a terminal status. */
  private async collectPendingRuns(ctx: PipelineContext, state: AnalysisState): Promise<void> {
    while (state.pendingRuns.size > 0) {
      await this.pollPendingRuns(ctx, state);

      if (state.pendingRuns.size > 0) {
        await this.delay(this.pollIntervalMs);
      }
    }
  }

  /** Executes one round-robin polling pass over the pending chunk runs. */
  private async pollPendingRuns(ctx: PipelineContext, state: AnalysisState): Promise<void> {
    for (const chunk of state.chunks) {
      const runId = state.pendingRuns.get(chunk.index);
      if (!runId) {
        continue;
      }

      await this.pollChunkRun(ctx, state, chunk, runId);
    }
  }

  /** Polls a single chunk run and captures its terminal outcome when available. */
  private async pollChunkRun(
    ctx: PipelineContext,
    state: AnalysisState,
    chunk: AnalyzeChunk,
    runId: string
  ): Promise<void> {
    ctx.emitter.emitProgress(
      state.submittedCount + state.completedCount,
      state.totalSteps,
      `Consultando resultado del fragmento ${chunk.index + 1} de ${state.chunks.length}...`
    );

    const pollResult = await ctx.analysisPort.pollChunkAnalysis(chunk.index, runId);
    console.log(this.getPollLogMessage(chunk.index, pollResult.status, pollResult.error));

    if (this.isPendingStatus(pollResult.status)) {
      return;
    }

    state.pendingRuns.delete(chunk.index);
    state.completedCount += 1;
    state.allSuggestions.push(...pollResult.suggestions);

    if (pollResult.error) {
      state.chunkErrors.push(pollResult.error);
    }

    ctx.emitter.emitProgress(
      state.submittedCount + state.completedCount,
      state.totalSteps,
      this.getPollCompletionMessage(chunk.index, pollResult.status)
    );
  }

  /** Copies collected analysis outputs back into the shared pipeline context. */
  private assignResults(ctx: PipelineContext, state: AnalysisState): void {
    ctx.rawSuggestions = state.allSuggestions;
    ctx.chunkErrors = state.chunkErrors;
  }

  /** Returns whether the pipeline must abort because no suggestions survived analysis. */
  private shouldAbort(state: AnalysisState): boolean {
    return state.allSuggestions.length === 0;
  }

  /** Aborts the pipeline with the user-facing reason derived from the analysis outcome. */
  private abortAnalysis(ctx: PipelineContext, errorCount: number): void {
    console.warn(`⚠️ [AnalyzeChunksHandler] Sin sugerencias. Errores de chunks: ${errorCount}`);

    ctx.aborted = true;
    ctx.abortReason =
      errorCount > 0
        ? `El análisis falló en ${errorCount} fragmento(s). Intenta de nuevo.`
        : "No se encontraron sugerencias editoriales.";
    ctx.emitter.emitAbort(ctx.abortReason);
  }

  /** Formats the submit log line for either successful queueing or submit failure. */
  private getSubmitLogMessage(chunkIndex: number, runId?: string, error?: string): string {
    if (runId) {
      return `🤖 [AnalyzeChunksHandler] Chunk ${chunkIndex + 1} enviado con runId "${runId}"`;
    }

    const errorDetail = error ? ` ⚠️ ${error}` : "";
    return `🤖 [AnalyzeChunksHandler] Chunk ${chunkIndex + 1} submit falló${errorDetail}`;
  }

  /** Formats the progress message emitted after each submit attempt. */
  private getSubmissionProgressMessage(chunkIndex: number, queued: boolean): string {
    return queued
      ? `Fragmento ${chunkIndex + 1} en cola. Esperando resultado...`
      : `No se pudo encolar el fragmento ${chunkIndex + 1}.`;
  }

  /** Formats the log line for an individual poll response. */
  private getPollLogMessage(chunkIndex: number, status: string, error?: string): string {
    const errorDetail = error ? ` ⚠️ ${error}` : "";
    return `🤖 [AnalyzeChunksHandler] Poll chunk ${chunkIndex + 1} → ${status}${errorDetail}`;
  }

  /** Formats the progress message emitted when a run reaches a terminal status. */
  private getPollCompletionMessage(chunkIndex: number, status: string): string {
    return status === "success"
      ? `Fragmento ${chunkIndex + 1} completado.`
      : `Fragmento ${chunkIndex + 1} terminó con error.`;
  }

  /** Identifies statuses that still require another polling cycle. */
  private isPendingStatus(status: string): boolean {
    return ["running", "pending", "waiting"].includes(status);
  }

  /** Waits between polling passes to avoid hammering the backend. */
  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
