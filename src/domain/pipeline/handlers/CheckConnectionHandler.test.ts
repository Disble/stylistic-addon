import type { IAnalysisPort, IDocumentPort } from "../../ports";
import type { PipelineContext } from "../PipelineContext";
import { PipelineEventEmitter } from "../PipelineEvents";
import { CheckConnectionHandler } from "./CheckConnectionHandler";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeContext(
  overrides: Partial<PipelineContext> = {},
): PipelineContext {
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
    genero: "general",
    maxChunkSize: 5000,
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("CheckConnectionHandler", () => {
  let handler: CheckConnectionHandler;
  let next: () => Promise<void>;

  beforeEach(() => {
    handler = new CheckConnectionHandler();
    next = vi.fn<() => Promise<void>>();
  });

  // -----------------------------------------------------------------------
  // Happy path — backend connected
  // -----------------------------------------------------------------------

  describe("happy path: backend connected", () => {
    it("calls next() when checkConnection returns true", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(true);

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });

    it("does not set aborted when connected", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(true);

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBeUndefined();
      expect(ctx.abortReason).toBeUndefined();
    });

    it("calls analysisPort.checkConnection exactly once", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(true);

      await handler.handle(ctx, next);

      expect(ctx.analysisPort.checkConnection).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Abort — backend not connected
  // -----------------------------------------------------------------------

  describe("abort: backend not connected", () => {
    it("aborts when checkConnection returns false", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(false);

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBe(true);
    });

    it("sets abort reason with backend unavailable message", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(false);

      await handler.handle(ctx, next);

      expect(ctx.abortReason).toBe(
        "Backend no disponible. Verifica que el servidor Mastra esté ejecutándose.",
      );
    });

    it("does NOT call next() when connection fails", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(false);

      await handler.handle(ctx, next);

      expect(next).not.toHaveBeenCalled();
    });
  });

  // -----------------------------------------------------------------------
  // Emitter events
  // -----------------------------------------------------------------------

  describe("emitter events", () => {
    it("emits phaseStart('connecting') at the beginning", async () => {
      const emitter = new PipelineEventEmitter();
      const onPhaseStart = vi.fn();
      emitter.subscribe({ onPhaseStart });

      const ctx = makeContext({ emitter });
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(true);

      await handler.handle(ctx, next);

      expect(onPhaseStart).toHaveBeenCalledWith(
        "connecting",
        "Conectando con el servidor...",
      );
    });

    it("emits phaseComplete('connecting') on success", async () => {
      const emitter = new PipelineEventEmitter();
      const onPhaseComplete = vi.fn();
      emitter.subscribe({ onPhaseComplete });

      const ctx = makeContext({ emitter });
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(true);

      await handler.handle(ctx, next);

      expect(onPhaseComplete).toHaveBeenCalledWith("connecting");
    });

    it("emits phaseStart but NOT phaseComplete when connection fails", async () => {
      const emitter = new PipelineEventEmitter();
      const onPhaseStart = vi.fn();
      const onPhaseComplete = vi.fn();
      emitter.subscribe({ onPhaseStart, onPhaseComplete });

      const ctx = makeContext({ emitter });
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(false);

      await handler.handle(ctx, next);

      expect(onPhaseStart).toHaveBeenCalledOnce();
      expect(onPhaseComplete).not.toHaveBeenCalled();
    });

    it("emits abort with reason when connection fails", async () => {
      const emitter = new PipelineEventEmitter();
      const onAbort = vi.fn();
      emitter.subscribe({ onAbort });

      const ctx = makeContext({ emitter });
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(false);

      await handler.handle(ctx, next);

      expect(onAbort).toHaveBeenCalledWith(
        "Backend no disponible. Verifica que el servidor Mastra esté ejecutándose.",
      );
    });

    it("event order on success: phaseStart before phaseComplete", async () => {
      const emitter = new PipelineEventEmitter();
      const timeline: string[] = [];
      emitter.subscribe({
        onPhaseStart: () => timeline.push("phaseStart"),
        onPhaseComplete: () => timeline.push("phaseComplete"),
      });

      const ctx = makeContext({ emitter });
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(true);

      await handler.handle(ctx, next);

      expect(timeline).toEqual(["phaseStart", "phaseComplete"]);
    });

    it("event order on failure: phaseStart before abort", async () => {
      const emitter = new PipelineEventEmitter();
      const timeline: string[] = [];
      emitter.subscribe({
        onPhaseStart: () => timeline.push("phaseStart"),
        onAbort: () => timeline.push("abort"),
      });

      const ctx = makeContext({ emitter });
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(false);

      await handler.handle(ctx, next);

      expect(timeline).toEqual(["phaseStart", "abort"]);
    });
  });

  // -----------------------------------------------------------------------
  // Error handling
  // -----------------------------------------------------------------------

  describe("error handling", () => {
    it("propagates errors thrown by analysisPort.checkConnection", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.analysisPort.checkConnection).mockRejectedValue(
        new Error("Network error"),
      );

      await expect(handler.handle(ctx, next)).rejects.toThrow("Network error");
      expect(next).not.toHaveBeenCalled();
    });
  });

  // -----------------------------------------------------------------------
  // Does not touch document port
  // -----------------------------------------------------------------------

  describe("isolation", () => {
    it("does not call any documentPort methods", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.analysisPort.checkConnection).mockResolvedValue(true);

      await handler.handle(ctx, next);

      expect(ctx.documentPort.getTextToAnalyze).not.toHaveBeenCalled();
      expect(ctx.documentPort.getAppliedOriginalTexts).not.toHaveBeenCalled();
      expect(ctx.documentPort.applySuggestions).not.toHaveBeenCalled();
      expect(ctx.documentPort.cleanupResolvedComments).not.toHaveBeenCalled();
    });
  });
});
