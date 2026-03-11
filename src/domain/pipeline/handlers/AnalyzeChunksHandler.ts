/* global console */

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

import { Suggestion } from "../../types";
import { PipelineContext } from "../PipelineContext";
import { PipelineHandler } from "./ReadTextHandler";

export class AnalyzeChunksHandler implements PipelineHandler {
  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    const scope = ctx.isSelection ? "selección" : "documento";
    const chunks = ctx.chunks!;
    console.log(
      `🤖 [AnalyzeChunksHandler] Fase 4: Analizando ${chunks.length} chunk(s) con Mastra...`
    );

    const allSuggestions: Suggestion[] = [];
    const chunkErrors: string[] = [];

    for (const chunk of chunks) {
      ctx.emitter.emitPhaseStart(
        "analyzing",
        `Analizando fragmento ${chunk.index + 1} de ${chunks.length} (${scope})...`
      );
      ctx.emitter.emitProgress(
        chunk.index + 1,
        chunks.length,
        `Analizando fragmento ${chunk.index + 1} de ${chunks.length}...`
      );

      console.log(
        `🤖 [AnalyzeChunksHandler] Enviando chunk ${chunk.index + 1}/${chunks.length} (${chunk.text.length} chars)`
      );
      const chunkResult = await ctx.analysisPort.analyzeChunk(chunk, ctx.profile, "es");
      console.log(
        `🤖 [AnalyzeChunksHandler] Chunk ${chunk.index + 1} → ${chunkResult.suggestions.length} sugerencia(s)${chunkResult.error ? " ⚠️ " + chunkResult.error : ""}`
      );

      allSuggestions.push(...chunkResult.suggestions);
      if (chunkResult.error) {
        chunkErrors.push(chunkResult.error);
      }
    }

    ctx.rawSuggestions = allSuggestions;
    ctx.chunkErrors = chunkErrors;

    if (allSuggestions.length === 0) {
      console.warn(
        `⚠️ [AnalyzeChunksHandler] Sin sugerencias. Errores de chunks: ${chunkErrors.length}`
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
}
