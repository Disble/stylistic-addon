/* global console */

/**
 * ChunkTextHandler — Phase 3 of the analysis pipeline.
 *
 * Splits the resolved text into chunks at paragraph boundaries.
 * Populates `ctx.chunks` for the analysis phase.
 *
 * Chunking is done on the frontend to:
 * - Respect the backend's context-window limits.
 * - Enable per-chunk retry (a failed chunk doesn't retry the whole document).
 * - Report granular progress to the user.
 *
 * @module ChunkTextHandler
 */

import { splitText } from "../../../infrastructure/chunker";
import type { PipelineContext } from "../PipelineContext";
import type { PipelineHandler } from "./ReadTextHandler.types";

/** Splits the resolved text into backend-sized chunks. */
export class ChunkTextHandler implements PipelineHandler {
  /** Reads the normalized text produced by ReadTextHandler. */
  private requireText(ctx: PipelineContext): string {
    if (!ctx.text) {
      throw new Error("ChunkTextHandler requires ctx.text before chunking.");
    }

    return ctx.text;
  }

  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    const scope = ctx.isSelection ? "selección" : "documento";
    console.log(`✂️ [ChunkTextHandler] Fase 3: Dividiendo ${scope} en chunks...`);

    const chunks = splitText(this.requireText(ctx), ctx.maxChunkSize);
    console.log(
      `✂️ [ChunkTextHandler] ${chunks.length} chunk(s) generados — genero: "${ctx.genero}"`
    );

    ctx.chunks = chunks;
    await next();
  }
}
