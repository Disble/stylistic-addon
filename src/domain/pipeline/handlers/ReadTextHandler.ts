/* global console */

/**
 * ReadTextHandler — Phase 1 of the analysis pipeline.
 *
 * Reads the text to analyze from the document. Resolves the active selection
 * first; falls back to the full document body if no selection is active.
 * Aborts the pipeline if the resolved text is empty.
 *
 * @module ReadTextHandler
 */

import type { PipelineContext } from "../PipelineContext";
import type { PipelineHandler } from "./ReadTextHandler.types";

/** Reads the text source that will enter the analysis pipeline. */
export class ReadTextHandler implements PipelineHandler {
  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    console.log("📖 [ReadTextHandler] Fase 1: Resolviendo texto a analizar...");
    ctx.emitter.emitPhaseStart("reading", "Leyendo texto...");

    const { text, isSelection } = await ctx.documentPort.getTextToAnalyze();
    const scope = isSelection ? "selección" : "documento";
    console.log(
      `📖 [ReadTextHandler] ${isSelection ? "Selección" : "Documento"} — ${text.length} chars`
    );

    if (!text || text.trim().length === 0) {
      console.warn("⚠️ [ReadTextHandler] Texto vacío — abortando pipeline");
      ctx.aborted = true;
      ctx.abortReason = "El documento está vacío. Escribe algo primero.";
      ctx.emitter.emitAbort(ctx.abortReason);
      return;
    }

    ctx.text = text;
    ctx.isSelection = isSelection;
    ctx.emitter.emitPhaseComplete("reading");
    console.log(`📖 [ReadTextHandler] OK — ${text.length} chars (${scope})`);

    await next();
  }
}
