/* global console */

/**
 * CheckConnectionHandler — Phase 2 of the analysis pipeline.
 *
 * Verifies that the analysis backend can analyze before starting analysis.
 * Fail-fast gate: aborts immediately with the backend's own reason (for
 * example, no model chosen) instead of letting the user wait through chunking
 * and analysis.
 *
 * @module CheckConnectionHandler
 */

import type { PipelineContext } from "../PipelineContext";
import type { PipelineHandler } from "./ReadTextHandler.types";

/** Verifies backend readiness before the costly analysis phases start. */
export class CheckConnectionHandler implements PipelineHandler {
  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    console.log("🔌 [CheckConnectionHandler] Fase 2: Verificando conexión...");
    ctx.emitter.emitPhaseStart("connecting", "Conectando con el servidor...");

    const readiness = await ctx.analysisPort.getReadiness();
    console.log(`🔌 [CheckConnectionHandler] Listo: ${readiness.ready ? "✅" : "❌"}`);

    if (!readiness.ready) {
      ctx.aborted = true;
      ctx.backendUnavailable = true;
      ctx.abortReason = readiness.reason ?? "El hub no puede analizar en este momento.";
      ctx.emitter.emitAbort(ctx.abortReason);
      return;
    }

    ctx.emitter.emitPhaseComplete("connecting");
    await next();
  }
}
