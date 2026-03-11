/**
 * Pipeline Orchestrator — assembles and executes the Chain of Responsibility.
 *
 * Receives an ordered list of `PipelineHandler` instances and runs them
 * sequentially. Each handler receives a `next` function to call the following
 * handler in the chain. If a handler sets `ctx.aborted = true` and does NOT
 * call `next()`, the chain stops gracefully.
 *
 * The orchestrator is agnostic to what handlers do — it only manages the
 * chain traversal and abort propagation. New analysis phases are added by
 * inserting a new handler into the array passed to the constructor.
 *
 * @module PipelineOrchestrator
 */

import { PipelineHandler } from "./handlers/ReadTextHandler";
import { PipelineContext } from "./PipelineContext";

/**
 * Executes a list of pipeline handlers in order, passing a shared context.
 *
 * Usage:
 * ```typescript
 * const orchestrator = new PipelineOrchestrator([
 *   new ReadTextHandler(),
 *   new CheckConnectionHandler(),
 *   new ChunkTextHandler(),
 *   new AnalyzeChunksHandler(),
 *   new DeduplicateHandler(),
 *   new GuardAppliedHandler(),
 *   new ApplySuggestionsHandler(),
 * ]);
 *
 * await orchestrator.run(ctx);
 * ```
 */
export class PipelineOrchestrator {
  constructor(private readonly handlers: PipelineHandler[]) {}

  /**
   * Runs the handler chain with the given context.
   *
   * Each handler is responsible for calling `next()` to continue the chain.
   * Handlers that set `ctx.aborted = true` should NOT call `next()`.
   *
   * @param ctx - The mutable pipeline context shared between all handlers.
   */
  async run(ctx: PipelineContext): Promise<void> {
    await this.execute(ctx, 0);
  }

  private async execute(ctx: PipelineContext, index: number): Promise<void> {
    if (index >= this.handlers.length) return;
    if (ctx.aborted) return;

    const handler = this.handlers[index];
    await handler.handle(ctx, () => this.execute(ctx, index + 1));
  }
}
