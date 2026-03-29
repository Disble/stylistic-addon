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

import { POLL_INTERVAL_MS } from "../../../infrastructure/config";
import type { Suggestion } from "../../types";
import type { PipelineContext } from "../PipelineContext";
import type { PipelineHandler } from "./ReadTextHandler";

export class AnalyzeChunksHandler implements PipelineHandler {
  constructor(private readonly pollIntervalMs: number = POLL_INTERVAL_MS) {}

  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    const scope = ctx.isSelection ? "selección" : "documento";
    const chunks = ctx.chunks!;
    console.log(
      `🤖 [AnalyzeChunksHandler] Fase 4: Analizando ${chunks.length} chunk(s) con Mastra...`,
    );

    const allSuggestions: Suggestion[] = [];
    const chunkErrors: string[] = [];
    const pendingRuns = new Map<number, string>();
    const totalSteps = Math.max(chunks.length * 2, 1);
    let submittedCount = 0;
    let completedCount = 0;

    for (const chunk of chunks) {
      ctx.emitter.emitPhaseStart(
        "analyzing",
        `Encolando fragmento ${chunk.index + 1} de ${chunks.length} (${scope})...`,
      );
      ctx.emitter.emitProgress(
        submittedCount,
        totalSteps,
        `Encolando fragmento ${chunk.index + 1} de ${chunks.length}...`,
      );

      console.log(
        `🤖 [AnalyzeChunksHandler] Encolando chunk ${chunk.index + 1}/${chunks.length} (${chunk.text.length} chars)`,
      );
      const submitResult = await ctx.analysisPort.submitChunkAnalysis(
        chunk,
        ctx.genero,
        "Disble",
      );
      submittedCount += 1;

      console.log(
        submitResult.runId
          ? `🤖 [AnalyzeChunksHandler] Chunk ${chunk.index + 1} enviado con runId "${submitResult.runId}"`
          : `🤖 [AnalyzeChunksHandler] Chunk ${chunk.index + 1} submit falló${submitResult.error ? " ⚠️ " + submitResult.error : ""}`,
      );

      ctx.emitter.emitProgress(
        submittedCount,
        totalSteps,
        submitResult.runId
          ? `Fragmento ${chunk.index + 1} en cola. Esperando resultado...`
          : `No se pudo encolar el fragmento ${chunk.index + 1}.`,
      );

      if (submitResult.runId) {
        pendingRuns.set(chunk.index, submitResult.runId);
        continue;
      }

      if (submitResult.error) {
        chunkErrors.push(submitResult.error);
      }
    }

    while (pendingRuns.size > 0) {
      for (const chunk of chunks) {
        const runId = pendingRuns.get(chunk.index);
        if (!runId) {
          continue;
        }

        ctx.emitter.emitProgress(
          submittedCount + completedCount,
          totalSteps,
          `Consultando resultado del fragmento ${chunk.index + 1} de ${chunks.length}...`,
        );

        const pollResult = await ctx.analysisPort.pollChunkAnalysis(
          chunk.index,
          runId,
        );
        console.log(
          `🤖 [AnalyzeChunksHandler] Poll chunk ${chunk.index + 1} → ${pollResult.status}${pollResult.error ? " ⚠️ " + pollResult.error : ""}`,
        );

        if (this.isPendingStatus(pollResult.status)) {
          continue;
        }

        pendingRuns.delete(chunk.index);
        completedCount += 1;
        allSuggestions.push(...pollResult.suggestions);

        if (pollResult.error) {
          chunkErrors.push(pollResult.error);
        }

        ctx.emitter.emitProgress(
          submittedCount + completedCount,
          totalSteps,
          pollResult.status === "success"
            ? `Fragmento ${chunk.index + 1} completado.`
            : `Fragmento ${chunk.index + 1} terminó con error.`,
        );
      }

      if (pendingRuns.size > 0) {
        await this.delay(this.pollIntervalMs);
      }
    }

    ctx.rawSuggestions = allSuggestions;
    ctx.chunkErrors = chunkErrors;

    if (allSuggestions.length === 0) {
      console.warn(
        `⚠️ [AnalyzeChunksHandler] Sin sugerencias. Errores de chunks: ${chunkErrors.length}`,
      );
      ctx.aborted = true;
      ctx.abortReason =
        chunkErrors.length > 0
          ? `El análisis falló en ${chunkErrors.length} fragmento(s). Intenta de nuevo.`
          : "No se encontraron sugerencias editoriales.";
      ctx.emitter.emitAbort(ctx.abortReason);
      return;
    }

    await next();
  }

  private isPendingStatus(status: string): boolean {
    return ["running", "pending", "waiting"].includes(status);
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
