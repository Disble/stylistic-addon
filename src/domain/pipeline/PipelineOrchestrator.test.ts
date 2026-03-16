import { PipelineOrchestrator } from "./PipelineOrchestrator";
import { PipelineContext } from "./PipelineContext";
import { PipelineEventEmitter } from "./PipelineEvents";
import type { PipelineHandler } from "./handlers/ReadTextHandler";
import type { IDocumentPort, IAnalysisPort } from "../ports";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Creates a minimal PipelineContext with mocked ports and a real emitter.
 * Override any field via the `overrides` parameter.
 */
function makeContext(overrides: Partial<PipelineContext> = {}): PipelineContext {
  const documentPort: IDocumentPort = {
    getTextToAnalyze: vi.fn(),
    getAppliedOriginalTexts: vi.fn(),
    applySuggestions: vi.fn(),
    cleanupResolvedComments: vi.fn(),
    acceptSuggestion: vi.fn(),
    rejectSuggestion: vi.fn(),
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
    profile: "general",
    maxChunkSize: 5000,
    ...overrides,
  };
}

/**
 * Creates a handler that records its invocation and calls `next()`.
 * Optionally runs a side-effect function before calling next.
 */
function makeHandler(
  name: string,
  calls: string[],
  sideEffect?: (ctx: PipelineContext) => void | Promise<void>
): PipelineHandler {
  return {
    handle: vi.fn(async (ctx: PipelineContext, next: () => Promise<void>) => {
      calls.push(name);
      if (sideEffect) await sideEffect(ctx);
      await next();
    }),
  };
}

/**
 * Creates a handler that records its invocation but does NOT call `next()`.
 */
function makeTerminalHandler(
  name: string,
  calls: string[],
  sideEffect?: (ctx: PipelineContext) => void | Promise<void>
): PipelineHandler {
  return {
    handle: vi.fn(async (ctx: PipelineContext, _next: () => Promise<void>) => {
      calls.push(name);
      if (sideEffect) await sideEffect(ctx);
      // Intentionally does NOT call next()
    }),
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("PipelineOrchestrator", () => {
  // -------------------------------------------------------------------------
  // Happy path — sequential handler execution
  // -------------------------------------------------------------------------

  describe("happy path", () => {
    it("executes all handlers in order when each calls next()", async () => {
      const calls: string[] = [];
      const handlers = [
        makeHandler("A", calls),
        makeHandler("B", calls),
        makeHandler("C", calls),
      ];
      const orchestrator = new PipelineOrchestrator(handlers);
      const ctx = makeContext();

      await orchestrator.run(ctx);

      expect(calls).toEqual(["A", "B", "C"]);
    });

    it("calls each handler's handle method exactly once", async () => {
      const calls: string[] = [];
      const handlers = [
        makeHandler("A", calls),
        makeHandler("B", calls),
      ];
      const orchestrator = new PipelineOrchestrator(handlers);

      await orchestrator.run(makeContext());

      for (const h of handlers) {
        expect(h.handle).toHaveBeenCalledOnce();
      }
    });

    it("passes the same context object to every handler", async () => {
      const receivedContexts: PipelineContext[] = [];
      const handlerA: PipelineHandler = {
        handle: vi.fn(async (ctx, next) => {
          receivedContexts.push(ctx);
          await next();
        }),
      };
      const handlerB: PipelineHandler = {
        handle: vi.fn(async (ctx, next) => {
          receivedContexts.push(ctx);
          await next();
        }),
      };
      const orchestrator = new PipelineOrchestrator([handlerA, handlerB]);
      const ctx = makeContext();

      await orchestrator.run(ctx);

      expect(receivedContexts).toHaveLength(2);
      expect(receivedContexts[0]).toBe(ctx);
      expect(receivedContexts[1]).toBe(ctx);
    });

    it("resolves the returned promise when all handlers complete", async () => {
      const calls: string[] = [];
      const handlers = [makeHandler("A", calls), makeHandler("B", calls)];
      const orchestrator = new PipelineOrchestrator(handlers);

      const result = orchestrator.run(makeContext());

      await expect(result).resolves.toBeUndefined();
    });
  });

  // -------------------------------------------------------------------------
  // Zero and single handler edge cases
  // -------------------------------------------------------------------------

  describe("edge cases: handler count", () => {
    it("completes immediately with zero handlers", async () => {
      const orchestrator = new PipelineOrchestrator([]);
      const ctx = makeContext();

      await expect(orchestrator.run(ctx)).resolves.toBeUndefined();
    });

    it("does not modify context with zero handlers", async () => {
      const orchestrator = new PipelineOrchestrator([]);
      const ctx = makeContext();

      await orchestrator.run(ctx);

      expect(ctx.aborted).toBeUndefined();
      expect(ctx.text).toBeUndefined();
    });

    it("executes a single handler correctly", async () => {
      const calls: string[] = [];
      const handler = makeHandler("only", calls);
      const orchestrator = new PipelineOrchestrator([handler]);

      await orchestrator.run(makeContext());

      expect(calls).toEqual(["only"]);
      expect(handler.handle).toHaveBeenCalledOnce();
    });

    it("single handler receives a working next() that resolves", async () => {
      // Verifying that calling next() with no more handlers doesn't throw
      let nextResolved = false;
      const handler: PipelineHandler = {
        handle: vi.fn(async (_ctx, next) => {
          await next();
          nextResolved = true;
        }),
      };
      const orchestrator = new PipelineOrchestrator([handler]);

      await orchestrator.run(makeContext());

      expect(nextResolved).toBe(true);
    });
  });

  // -------------------------------------------------------------------------
  // Context mutation across handlers
  // -------------------------------------------------------------------------

  describe("context passing between handlers", () => {
    it("allows handlers to mutate context for subsequent handlers", async () => {
      const calls: string[] = [];

      const writerHandler = makeHandler("writer", calls, (ctx) => {
        ctx.text = "hello world";
        ctx.isSelection = false;
      });

      let capturedText: string | undefined;
      let capturedIsSelection: boolean | undefined;
      const readerHandler: PipelineHandler = {
        handle: vi.fn(async (ctx, next) => {
          calls.push("reader");
          capturedText = ctx.text;
          capturedIsSelection = ctx.isSelection;
          await next();
        }),
      };

      const orchestrator = new PipelineOrchestrator([writerHandler, readerHandler]);
      await orchestrator.run(makeContext());

      expect(calls).toEqual(["writer", "reader"]);
      expect(capturedText).toBe("hello world");
      expect(capturedIsSelection).toBe(false);
    });

    it("allows multiple handlers to progressively build up context", async () => {
      const calls: string[] = [];

      const h1 = makeHandler("setText", calls, (ctx) => {
        ctx.text = "some text";
      });

      const h2 = makeHandler("setChunks", calls, (ctx) => {
        ctx.chunks = [{ text: ctx.text!, index: 0, total: 1, startOffset: 0 }];
      });

      let finalChunks: unknown;
      const h3: PipelineHandler = {
        handle: vi.fn(async (ctx, next) => {
          calls.push("readChunks");
          finalChunks = ctx.chunks;
          await next();
        }),
      };

      const orchestrator = new PipelineOrchestrator([h1, h2, h3]);
      await orchestrator.run(makeContext());

      expect(calls).toEqual(["setText", "setChunks", "readChunks"]);
      expect(finalChunks).toEqual([
        { text: "some text", index: 0, total: 1, startOffset: 0 },
      ]);
    });
  });

  // -------------------------------------------------------------------------
  // Abort mechanism
  // -------------------------------------------------------------------------

  describe("abort mechanism", () => {
    it("skips subsequent handlers when ctx.aborted is set before calling next()", async () => {
      const calls: string[] = [];

      const aborter = makeHandler("aborter", calls, (ctx) => {
        ctx.aborted = true;
        ctx.abortReason = "empty document";
        // Note: this handler still calls next() (via makeHandler),
        // but the orchestrator checks ctx.aborted before executing the next handler
      });

      const skipped = makeHandler("skipped", calls);
      const orchestrator = new PipelineOrchestrator([aborter, skipped]);

      await orchestrator.run(makeContext());

      expect(calls).toEqual(["aborter"]);
      expect(skipped.handle).not.toHaveBeenCalled();
    });

    it("skips all remaining handlers after abort, not just the next one", async () => {
      const calls: string[] = [];

      const h1 = makeHandler("h1", calls);
      const aborter = makeHandler("aborter", calls, (ctx) => {
        ctx.aborted = true;
      });
      const h3 = makeHandler("h3", calls);
      const h4 = makeHandler("h4", calls);

      const orchestrator = new PipelineOrchestrator([h1, aborter, h3, h4]);
      await orchestrator.run(makeContext());

      expect(calls).toEqual(["h1", "aborter"]);
    });

    it("does not execute any handler when context starts already aborted", async () => {
      const calls: string[] = [];
      const handler = makeHandler("never", calls);
      const orchestrator = new PipelineOrchestrator([handler]);

      const ctx = makeContext({ aborted: true });
      await orchestrator.run(ctx);

      expect(calls).toEqual([]);
      expect(handler.handle).not.toHaveBeenCalled();
    });

    it("preserves the abort reason on the context", async () => {
      const calls: string[] = [];
      const aborter = makeHandler("aborter", calls, (ctx) => {
        ctx.aborted = true;
        ctx.abortReason = "Backend unreachable";
      });
      const orchestrator = new PipelineOrchestrator([aborter, makeHandler("skipped", calls)]);
      const ctx = makeContext();

      await orchestrator.run(ctx);

      expect(ctx.aborted).toBe(true);
      expect(ctx.abortReason).toBe("Backend unreachable");
    });

    it("a handler that does NOT call next() also stops the chain (terminal handler)", async () => {
      const calls: string[] = [];

      const terminal = makeTerminalHandler("terminal", calls);
      const afterTerminal = makeHandler("afterTerminal", calls);

      const orchestrator = new PipelineOrchestrator([terminal, afterTerminal]);
      await orchestrator.run(makeContext());

      expect(calls).toEqual(["terminal"]);
      expect(afterTerminal.handle).not.toHaveBeenCalled();
    });

    it("abort in first handler skips all remaining handlers", async () => {
      const calls: string[] = [];

      const firstAborter = makeHandler("first", calls, (ctx) => {
        ctx.aborted = true;
      });

      const orchestrator = new PipelineOrchestrator([
        firstAborter,
        makeHandler("second", calls),
        makeHandler("third", calls),
      ]);

      await orchestrator.run(makeContext());

      expect(calls).toEqual(["first"]);
    });
  });

  // -------------------------------------------------------------------------
  // Error handling
  // -------------------------------------------------------------------------

  describe("error handling", () => {
    it("propagates errors thrown by a handler", async () => {
      const failing: PipelineHandler = {
        handle: vi.fn(async () => {
          throw new Error("handler exploded");
        }),
      };
      const orchestrator = new PipelineOrchestrator([failing]);

      await expect(orchestrator.run(makeContext())).rejects.toThrow("handler exploded");
    });

    it("does not execute handlers after the one that threw", async () => {
      const calls: string[] = [];

      const failing: PipelineHandler = {
        handle: vi.fn(async () => {
          calls.push("failing");
          throw new Error("boom");
        }),
      };
      const afterFailing = makeHandler("afterFailing", calls);

      const orchestrator = new PipelineOrchestrator([failing, afterFailing]);

      await expect(orchestrator.run(makeContext())).rejects.toThrow("boom");
      expect(calls).toEqual(["failing"]);
      expect(afterFailing.handle).not.toHaveBeenCalled();
    });

    it("runs handlers before the failing one successfully", async () => {
      const calls: string[] = [];

      const ok = makeHandler("ok", calls);
      const failing: PipelineHandler = {
        handle: vi.fn(async () => {
          calls.push("failing");
          throw new Error("boom");
        }),
      };

      const orchestrator = new PipelineOrchestrator([ok, failing]);

      await expect(orchestrator.run(makeContext())).rejects.toThrow();
      expect(calls).toEqual(["ok", "failing"]);
    });

    it("propagates the original error type (not wrapped)", async () => {
      class CustomError extends Error {
        constructor(public code: number) {
          super("custom");
        }
      }

      const failing: PipelineHandler = {
        handle: vi.fn(async () => {
          throw new CustomError(42);
        }),
      };
      const orchestrator = new PipelineOrchestrator([failing]);

      try {
        await orchestrator.run(makeContext());
        expect.unreachable("should have thrown");
      } catch (err) {
        expect(err).toBeInstanceOf(CustomError);
        expect((err as CustomError).code).toBe(42);
      }
    });

    it("propagates non-Error throws (e.g., strings)", async () => {
      const failing: PipelineHandler = {
        handle: vi.fn(async () => {
          throw "string error";
        }),
      };
      const orchestrator = new PipelineOrchestrator([failing]);

      try {
        await orchestrator.run(makeContext());
        expect.unreachable("should have thrown");
      } catch (err) {
        expect(err).toBe("string error");
      }
    });
  });

  // -------------------------------------------------------------------------
  // Chain of Responsibility — next() behavior
  // -------------------------------------------------------------------------

  describe("chain of responsibility (next() behavior)", () => {
    it("handler can perform work after calling next() (post-processing)", async () => {
      const timeline: string[] = [];

      const wrapper: PipelineHandler = {
        handle: vi.fn(async (_ctx, next) => {
          timeline.push("wrapper:before");
          await next();
          timeline.push("wrapper:after");
        }),
      };

      const inner: PipelineHandler = {
        handle: vi.fn(async (_ctx, next) => {
          timeline.push("inner");
          await next();
        }),
      };

      const orchestrator = new PipelineOrchestrator([wrapper, inner]);
      await orchestrator.run(makeContext());

      expect(timeline).toEqual(["wrapper:before", "inner", "wrapper:after"]);
    });

    it("calling next() twice from the same handler rejects and does not re-run the chain", async () => {
      const calls: string[] = [];

      const doubleNext: PipelineHandler = {
        handle: vi.fn(async (_ctx, next) => {
          calls.push("double");
          await next();
          await next();
        }),
      };

      const counter = makeHandler("counted", calls);
      const orchestrator = new PipelineOrchestrator([doubleNext, counter]);

      await expect(orchestrator.run(makeContext())).rejects.toThrow(
        "next() called multiple times in the same handler"
      );
      expect(calls).toEqual(["double", "counted"]);
      expect(counter.handle).toHaveBeenCalledOnce();
    });

    it("not calling next() stops the chain even without setting aborted", async () => {
      const calls: string[] = [];

      const stopper: PipelineHandler = {
        handle: vi.fn(async () => {
          calls.push("stopper");
          // Does not call next() and does not set ctx.aborted
        }),
      };

      const unreached = makeHandler("unreached", calls);
      const orchestrator = new PipelineOrchestrator([stopper, unreached]);
      await orchestrator.run(makeContext());

      expect(calls).toEqual(["stopper"]);
    });
  });

  // -------------------------------------------------------------------------
  // Interaction with PipelineContext's emitter
  // -------------------------------------------------------------------------

  describe("interaction with emitter (via handlers)", () => {
    it("handlers can emit events through the context emitter", async () => {
      const onPhaseStart = vi.fn();
      const emitter = new PipelineEventEmitter();
      emitter.subscribe({ onPhaseStart });

      const calls: string[] = [];
      const emitting = makeHandler("emitting", calls, (ctx) => {
        ctx.emitter.emitPhaseStart("reading", "Leyendo texto...");
      });

      const orchestrator = new PipelineOrchestrator([emitting]);
      await orchestrator.run(makeContext({ emitter }));

      expect(onPhaseStart).toHaveBeenCalledWith("reading", "Leyendo texto...");
    });

    it("abort emits event when handler uses emitter before aborting", async () => {
      const onAbort = vi.fn();
      const emitter = new PipelineEventEmitter();
      emitter.subscribe({ onAbort });

      const calls: string[] = [];
      const aborting = makeHandler("aborting", calls, (ctx) => {
        ctx.aborted = true;
        ctx.abortReason = "No text found";
        ctx.emitter.emitAbort("No text found");
      });

      const orchestrator = new PipelineOrchestrator([aborting]);
      await orchestrator.run(makeContext({ emitter }));

      expect(onAbort).toHaveBeenCalledWith("No text found");
    });
  });

  // -------------------------------------------------------------------------
  // Async behavior
  // -------------------------------------------------------------------------

  describe("async behavior", () => {
    it("waits for each handler to complete before calling the next", async () => {
      const timeline: string[] = [];

      const slow: PipelineHandler = {
        handle: vi.fn(async (_ctx, next) => {
          timeline.push("slow:start");
          await new Promise((r) => setTimeout(r, 10));
          timeline.push("slow:end");
          await next();
        }),
      };

      const fast: PipelineHandler = {
        handle: vi.fn(async (_ctx, next) => {
          timeline.push("fast");
          await next();
        }),
      };

      const orchestrator = new PipelineOrchestrator([slow, fast]);
      await orchestrator.run(makeContext());

      expect(timeline).toEqual(["slow:start", "slow:end", "fast"]);
    });

    it("handles rejected promises from async handlers", async () => {
      const asyncFail: PipelineHandler = {
        handle: vi.fn(async () => {
          await new Promise((_, reject) => setTimeout(() => reject(new Error("async boom")), 1));
        }),
      };

      const orchestrator = new PipelineOrchestrator([asyncFail]);

      await expect(orchestrator.run(makeContext())).rejects.toThrow("async boom");
    });
  });

  // -------------------------------------------------------------------------
  // Immutability of handler list
  // -------------------------------------------------------------------------

  describe("handler list", () => {
    it("uses the handlers provided at construction time", async () => {
      const calls: string[] = [];
      const handlers = [makeHandler("A", calls), makeHandler("B", calls)];
      const orchestrator = new PipelineOrchestrator(handlers);

      await orchestrator.run(makeContext());

      expect(calls).toEqual(["A", "B"]);
    });

    it("can be run multiple times with different contexts", async () => {
      const calls: string[] = [];
      const handlers = [makeHandler("A", calls)];
      const orchestrator = new PipelineOrchestrator(handlers);

      await orchestrator.run(makeContext());
      await orchestrator.run(makeContext());

      expect(calls).toEqual(["A", "A"]);
    });

    it("each run is independent — abort in first run does not affect second", async () => {
      const calls: string[] = [];

      const conditionalHandler: PipelineHandler = {
        handle: vi.fn(async (ctx, next) => {
          calls.push("handler");
          await next();
        }),
      };

      const orchestrator = new PipelineOrchestrator([conditionalHandler]);

      // First run: aborted context
      const ctx1 = makeContext({ aborted: true });
      await orchestrator.run(ctx1);

      // Second run: normal context
      const ctx2 = makeContext();
      await orchestrator.run(ctx2);

      // First run skipped, second run executed
      expect(calls).toEqual(["handler"]);
    });
  });

  // -------------------------------------------------------------------------
  // Realistic scenario: mimicking the real pipeline
  // -------------------------------------------------------------------------

  describe("realistic pipeline scenario", () => {
    it("simulates a mini pipeline: read → chunk → analyze → apply", async () => {
      const calls: string[] = [];

      const readHandler = makeHandler("read", calls, (ctx) => {
        ctx.text = "Hello world. This is a test.";
        ctx.isSelection = false;
      });

      const chunkHandler = makeHandler("chunk", calls, (ctx) => {
        ctx.chunks = [
          { text: ctx.text!, index: 0, total: 1, startOffset: 0 },
        ];
      });

      const analyzeHandler = makeHandler("analyze", calls, (ctx) => {
        ctx.rawSuggestions = [
          {
            id: "s1",
            originalText: "Hello world",
            suggestedText: "Hello, world",
            justification: "Missing comma",
            category: "Puntuación",
            severity: "low",
          },
        ];
      });

      const applyHandler = makeHandler("apply", calls, (ctx) => {
        ctx.result = {
          successCount: ctx.rawSuggestions?.length ?? 0,
          failedSuggestions: [],
        };
      });

      const orchestrator = new PipelineOrchestrator([
        readHandler,
        chunkHandler,
        analyzeHandler,
        applyHandler,
      ]);
      const ctx = makeContext();

      await orchestrator.run(ctx);

      expect(calls).toEqual(["read", "chunk", "analyze", "apply"]);
      expect(ctx.text).toBe("Hello world. This is a test.");
      expect(ctx.chunks).toHaveLength(1);
      expect(ctx.rawSuggestions).toHaveLength(1);
      expect(ctx.result).toEqual({ successCount: 1, failedSuggestions: [] });
    });

    it("simulates abort at read phase (empty document)", async () => {
      const calls: string[] = [];

      const readHandler = makeHandler("read", calls, (ctx) => {
        // Simulate empty document — abort
        ctx.aborted = true;
        ctx.abortReason = "El documento está vacío. Escribe algo primero.";
        ctx.emitter.emitAbort(ctx.abortReason!);
      });

      const chunkHandler = makeHandler("chunk", calls);
      const analyzeHandler = makeHandler("analyze", calls);

      const onAbort = vi.fn();
      const emitter = new PipelineEventEmitter();
      emitter.subscribe({ onAbort });

      const orchestrator = new PipelineOrchestrator([
        readHandler,
        chunkHandler,
        analyzeHandler,
      ]);
      const ctx = makeContext({ emitter });

      await orchestrator.run(ctx);

      expect(calls).toEqual(["read"]);
      expect(ctx.aborted).toBe(true);
      expect(ctx.abortReason).toBe("El documento está vacío. Escribe algo primero.");
      expect(onAbort).toHaveBeenCalledWith("El documento está vacío. Escribe algo primero.");
    });

    it("simulates error at analyze phase", async () => {
      const calls: string[] = [];

      const readHandler = makeHandler("read", calls, (ctx) => {
        ctx.text = "Some text";
      });

      const analyzeHandler: PipelineHandler = {
        handle: vi.fn(async () => {
          calls.push("analyze");
          throw new Error("Backend timeout");
        }),
      };

      const applyHandler = makeHandler("apply", calls);

      const orchestrator = new PipelineOrchestrator([
        readHandler,
        analyzeHandler,
        applyHandler,
      ]);
      const ctx = makeContext();

      await expect(orchestrator.run(ctx)).rejects.toThrow("Backend timeout");
      expect(calls).toEqual(["read", "analyze"]);
      expect(ctx.text).toBe("Some text");
      expect(applyHandler.handle).not.toHaveBeenCalled();
    });
  });
});
