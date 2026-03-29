/* global console */

/**
 * ApplySuggestionsHandler — Phase 6 of the analysis pipeline.
 *
 * Applies pending suggestions as native Word tracked changes with embedded
 * justification comments. Delegates to `IDocumentPort.applySuggestions()`.
 *
 * Progress is reported per suggestion via `ProgressCallback`, which maps
 * to `PipelineEventEmitter.emitProgress()` for the UI.
 *
 * After completion, emits the final `onComplete` event with all results.
 *
 * @module ApplySuggestionsHandler
 */

import type { AnalysisPhase, ProgressCallback } from "../../types";
import type { PipelineContext } from "../PipelineContext";
import type { PipelineHandler } from "./ReadTextHandler";

export class ApplySuggestionsHandler implements PipelineHandler {
  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    const pending = ctx.pendingSuggestions!;
    console.log(
      `📝 [ApplySuggestionsHandler] Fase 6: Aplicando ${pending.length} sugerencias como Track Changes...`,
    );

    // Bridge: ProgressCallback → PipelineEventEmitter
    const onProgress: ProgressCallback = (
      _phase: AnalysisPhase,
      current: number,
      total: number,
      message: string,
    ) => {
      ctx.emitter.emitProgress(current, total, message);
    };

    const result = await ctx.documentPort.applySuggestions(pending, onProgress);
    console.log(
      `📝 [ApplySuggestionsHandler] Resultado: ${result.successCount} aplicadas, ${result.failedSuggestions.length} fallidas`,
    );

    ctx.result = result;
    ctx.emitter.emitPhaseComplete("applying");
    ctx.emitter.emitComplete(
      pending,
      result,
      ctx.chunkErrors ?? [],
      ctx.isSelection ?? false,
    );

    await next();
  }
}
