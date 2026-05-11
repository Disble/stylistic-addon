/* global console */

/**
 * CheckConnectionHandler — Phase 2 of the analysis pipeline.
 *
 * Verifies that the Mastra backend is reachable before starting analysis.
 * Fail-fast gate: aborts immediately if the backend is unavailable,
 * saving the user from waiting through chunking and analysis.
 *
 * @module CheckConnectionHandler
 */

import { MASTRA_POLL_BYPASS_ENABLED } from "../../../infrastructure/config";
import type { PipelineContext } from "../PipelineContext";
import type { PipelineHandler } from "./ReadTextHandler.types";

/** Verifies backend reachability before the costly analysis phases start. */
export class CheckConnectionHandler implements PipelineHandler {
  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    console.log("🔌 [CheckConnectionHandler] Fase 2: Verificando conexión...");

    if (MASTRA_POLL_BYPASS_ENABLED) {
      console.log(
        "🧪 [CheckConnectionHandler] Bypass activo: se omite la conexión real al backend"
      );
      ctx.emitter.emitPhaseComplete("connecting");
      await next();
      return;
    }

    ctx.emitter.emitPhaseStart("connecting", "Conectando con el servidor...");

    const connected = await ctx.analysisPort.checkConnection();
    console.log(`🔌 [CheckConnectionHandler] Conexión: ${connected ? "✅ OK" : "❌ FALLO"}`);

    if (!connected) {
      ctx.aborted = true;
      ctx.abortReason = "Backend no disponible. Verifica que el servidor Mastra esté ejecutándose.";
      ctx.emitter.emitAbort(ctx.abortReason);
      return;
    }

    ctx.emitter.emitPhaseComplete("connecting");
    await next();
  }
}
