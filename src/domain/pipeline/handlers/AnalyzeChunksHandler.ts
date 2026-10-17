/* global console */

/**
 * AnalyzeChunksHandler — Phase 4 of the analysis pipeline.
 *
 * Starts one backend analysis covering every text chunk, polls it until it
 * finishes, and collects the suggestions. Applies the Partial Success
 * philosophy: if some chunks fail, suggestions from successful chunks are still
 * applied, and each failure carries the backend's real cause.
 *
 * Progress is emitted via the PipelineEventEmitter as chunks finish.
 *
 * @module AnalyzeChunksHandler
 */

import { POLL_INTERVAL_MS } from "../../../infrastructure/config";
import {
  collectAnalysisSuggestions,
  countDroppedSuggestions,
  countFinishedChunks,
  describeAnalysisFailures,
  describeNoSuggestionsReason,
} from "../../analysis/Analysis.helpers";
import type { AnalysisSnapshot } from "../../analysis/Analysis.types";
import { AnalysisPoller } from "../../analysis/AnalysisPoller";
import type { TextChunk } from "../../chunking/TextChunk.types";
import type { PipelineContext } from "../PipelineContext";
import type { PipelineHandler } from "./ReadTextHandler.types";

/** Starts and polls one analysis for the whole analysis phase. */
export class AnalyzeChunksHandler implements PipelineHandler {
  constructor(private readonly pollIntervalMs: number = POLL_INTERVAL_MS) {}

  /** Starts the analysis, waits for it to finish, and hands its suggestions on. */
  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    const chunks = this.requireChunks(ctx);
    const scope = ctx.isSelection ? "selección" : "documento";

    console.log(`🤖 [AnalyzeChunksHandler] Fase 4: Analizando ${chunks.length} chunk(s)...`);
    ctx.emitter.emitPhaseStart("analyzing", `Analizando ${scope}...`);

    const started = await this.startAnalysis(ctx, chunks);
    if (!started) {
      return;
    }

    ctx.activeAnalysisId = started.analysisId;
    ctx.emitter.emitProgress(0, chunks.length, `Analizando ${chunks.length} fragmento(s)...`);

    const poller = new AnalysisPoller(ctx.analysisPort, this.pollIntervalMs);
    const outcome = await poller.waitUntilFinished(started, {
      shouldCancel: ctx.shouldCancelAnalysis,
      onPoll: (snapshot) => {
        const finished = countFinishedChunks(snapshot);
        console.log(
          `🤖 [AnalyzeChunksHandler] Poll → ${snapshot.status} (${finished}/${chunks.length})`
        );
        ctx.emitter.emitProgress(
          finished,
          chunks.length,
          `Fragmentos terminados: ${finished} de ${chunks.length}.`
        );
      },
    });
    ctx.activeAnalysisId = undefined;

    if (outcome.kind === "poll-failed") {
      ctx.retryableAnalysisId = started.analysisId;
      this.abort(
        ctx,
        `La consulta del análisis falló localmente: ${outcome.message}. Reintentá la consulta.`
      );
      return;
    }

    if (outcome.kind === "canceled-locally" || outcome.snapshot.status === "canceled") {
      ctx.chunkErrors = [];
      this.abort(ctx, "Análisis cancelado por el usuario.");
      return;
    }

    this.collectResults(ctx, outcome.snapshot);

    if ((ctx.rawSuggestions ?? []).length === 0) {
      this.abort(ctx, describeNoSuggestionsReason(ctx.chunkErrors ?? []));
      return;
    }

    await next();
  }

  /** Reads pipeline chunks after the previous handler populated them. */
  private requireChunks(ctx: PipelineContext): TextChunk[] {
    if (!ctx.chunks) {
      throw new Error("AnalyzeChunksHandler requires ctx.chunks before analysis starts.");
    }

    return ctx.chunks;
  }

  /** Reads the stable document UUID required to start an analysis. */
  private requireDocumentUuid(ctx: PipelineContext): string {
    if (!ctx.documentUuid) {
      throw new Error("AnalyzeChunksHandler requires ctx.documentUuid before analysis starts.");
    }

    return ctx.documentUuid;
  }

  /** Starts the analysis, or aborts with the backend's explanation when it refuses. */
  private async startAnalysis(
    ctx: PipelineContext,
    chunks: TextChunk[]
  ): Promise<AnalysisSnapshot | undefined> {
    const documentUuid = this.requireDocumentUuid(ctx);

    try {
      const snapshot = await ctx.analysisPort.startAnalysis({
        documentUuid,
        genre: ctx.genero,
        chunks,
      });
      console.log(`🤖 [AnalyzeChunksHandler] Análisis iniciado: "${snapshot.analysisId}"`);
      return snapshot;
    } catch (error) {
      const message =
        error instanceof Error && error.message.trim().length > 0
          ? error.message.trim()
          : "error desconocido";
      ctx.chunkErrors = [message];
      this.abort(ctx, `No se pudo iniciar el análisis: ${message}`);
      return undefined;
    }
  }

  /** Copies suggestions and per-chunk failures of a finished analysis into the context. */
  private collectResults(ctx: PipelineContext, snapshot: AnalysisSnapshot): void {
    ctx.rawSuggestions = collectAnalysisSuggestions(snapshot);
    ctx.chunkErrors = describeAnalysisFailures(snapshot);

    const dropped = countDroppedSuggestions(snapshot);
    if (dropped > 0) {
      console.warn(`⚠️ [AnalyzeChunksHandler] El backend descartó ${dropped} sugerencia(s).`);
    }
  }

  /** Stops the pipeline with a user-facing reason. */
  private abort(ctx: PipelineContext, reason: string): void {
    console.warn(`⚠️ [AnalyzeChunksHandler] ${reason}`);
    ctx.aborted = true;
    ctx.abortReason = reason;
    ctx.emitter.emitAbort(reason);
  }
}
