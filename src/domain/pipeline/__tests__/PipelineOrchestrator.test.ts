import { describe, expect, it, vi } from "vitest";
import type { IAnalysisPort, IDocumentPort } from "../../ports";
import type { PipelineContext } from "../PipelineContext";
import { PipelineEventEmitter } from "../PipelineEvents";
import type { PipelineHandler } from "../handlers/ReadTextHandler.types";
import { PipelineOrchestrator } from "../PipelineOrchestrator";

/** Creates a minimal pipeline context with mocked ports. */
function makeContext(overrides: Partial<PipelineContext> = {}): PipelineContext {
  const documentPort: IDocumentPort = {
    getTextToAnalyze: vi.fn(),
    getDocumentUuid: vi.fn(),
    getAppliedOriginalTexts: vi.fn(),
    applySuggestions: vi.fn(),
    getCleanupPreview: vi.fn(),
    cleanupResolvedComments: vi.fn(),
    acceptSuggestion: vi.fn(),
    rejectSuggestion: vi.fn(),
    getDocumentReviewState: vi.fn(),
    disableTrackChanges: vi.fn(),
    navigateToText: vi.fn(),
    subscribeSelectionChanges: vi.fn(() => () => {}),
  };

  const analysisPort: IAnalysisPort = {
    checkConnection: vi.fn(),
    submitChunkAnalysis: vi.fn(),
    pollChunkAnalysis: vi.fn(),
  };

  return {
    documentPort,
    analysisPort,
    emitter: new PipelineEventEmitter(),
    genero: "general",
    maxChunkSize: 1000,
    ...overrides,
  };
}

/** Creates a handler that records execution and optionally mutates shared context. */
function makeHandler(
  name: string,
  calls: string[],
  sideEffect?: (ctx: PipelineContext) => void | Promise<void>
): PipelineHandler {
  return {
    handle: vi.fn(async (ctx: PipelineContext, next: () => Promise<void>) => {
      calls.push(name);
      await sideEffect?.(ctx);
      await next();
    }),
  };
}

describe("PipelineOrchestrator", () => {
  it("executes handlers in order and shares the same mutable context", async () => {
    const calls: string[] = [];
    const seenTexts: Array<string | undefined> = [];
    const orchestrator = new PipelineOrchestrator([
      makeHandler("read", calls, (ctx) => {
        ctx.text = "hola mundo";
      }),
      {
        handle: vi.fn(async (ctx, next) => {
          calls.push("consume");
          seenTexts.push(ctx.text);
          await next();
        }),
      },
    ]);

    const ctx = makeContext();
    await orchestrator.run(ctx);

    expect(calls).toEqual(["read", "consume"]);
    expect(seenTexts).toEqual(["hola mundo"]);
  });

  it("stops the chain when a handler does not call next()", async () => {
    const calls: string[] = [];
    const orchestrator = new PipelineOrchestrator([
      {
        handle: vi.fn(async (_ctx, _next) => {
          calls.push("terminal");
        }),
      },
      makeHandler("should-not-run", calls),
    ]);

    await orchestrator.run(makeContext());

    expect(calls).toEqual(["terminal"]);
  });

  it("stops before the next handler when the context is aborted", async () => {
    const calls: string[] = [];
    const orchestrator = new PipelineOrchestrator([
      makeHandler("aborter", calls, (ctx) => {
        ctx.aborted = true;
      }),
      makeHandler("should-not-run", calls),
    ]);

    await orchestrator.run(makeContext());

    expect(calls).toEqual(["aborter"]);
  });

  it("throws when one handler calls next() twice", async () => {
    const orchestrator = new PipelineOrchestrator([
      {
        handle: vi.fn(async (_ctx, next) => {
          await next();
          await next();
        }),
      },
    ]);

    await expect(orchestrator.run(makeContext())).rejects.toThrow(
      "next() called multiple times in the same handler"
    );
  });
});
