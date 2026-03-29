/* global console */

/**
 * GuardAppliedHandler — Phase 5b of the analysis pipeline.
 *
 * Guard pattern: filters out suggestions whose `originalText` is already
 * present as a Stylistic tracked deletion in the document. Prevents
 * duplicate tracked changes when the user re-runs analysis on a document
 * that still has pending Stylistic changes.
 *
 * Why needed: a `<w:del>` node's text is still present in the document body
 * and searchable via `body.search()`. Without this guard, the second run
 * would find the already-deleted text and wrap it in a second tracked change,
 * corrupting the document.
 *
 * @module GuardAppliedHandler
 */

import type { PipelineContext } from "../PipelineContext";
import type { PipelineHandler } from "./ReadTextHandler";

export class GuardAppliedHandler implements PipelineHandler {
  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    console.log(
      "🛡️ [GuardAppliedHandler] Fase 5b: Verificando sugerencias ya aplicadas...",
    );

    const unique = ctx.uniqueSuggestions!;
    const appliedTexts = await ctx.documentPort.getAppliedOriginalTexts();
    const pending = unique.filter((s) => !appliedTexts.has(s.originalText));
    const skipped = unique.length - pending.length;

    if (skipped > 0) {
      console.log(
        `🛡️ [GuardAppliedHandler] ${skipped} sugerencia(s) ya aplicada(s) — omitidas`,
      );
    }

    if (pending.length === 0) {
      ctx.aborted = true;
      ctx.abortReason =
        skipped > 0
          ? "Todas las sugerencias ya están aplicadas en el documento."
          : "No se encontraron sugerencias editoriales.";
      ctx.emitter.emitAbort(ctx.abortReason);
      return;
    }

    ctx.pendingSuggestions = pending;
    await next();
  }
}
