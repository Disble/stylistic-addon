import { ReadTextHandler } from "./ReadTextHandler";
import { PipelineContext } from "../PipelineContext";
import { PipelineEventEmitter } from "../PipelineEvents";
import type { IDocumentPort, IAnalysisPort } from "../../ports";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeContext(overrides: Partial<PipelineContext> = {}): PipelineContext {
  const documentPort: IDocumentPort = {
    getTextToAnalyze: vi.fn(),
    getAppliedOriginalTexts: vi.fn(),
    applySuggestions: vi.fn(),
    cleanupResolvedComments: vi.fn(),
  };

  const analysisPort: IAnalysisPort = {
    checkConnection: vi.fn(),
    analyzeChunk: vi.fn(),
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

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("ReadTextHandler", () => {
  let handler: ReadTextHandler;
  let next: () => Promise<void>;

  beforeEach(() => {
    handler = new ReadTextHandler();
    next = vi.fn<() => Promise<void>>();
  });

  // -----------------------------------------------------------------------
  // Happy path — document text
  // -----------------------------------------------------------------------

  describe("happy path: full document text", () => {
    it("reads text from documentPort and sets ctx.text", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "Hello world",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(ctx.text).toBe("Hello world");
    });

    it("sets ctx.isSelection to false when reading full document", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "Some document text",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(ctx.isSelection).toBe(false);
    });

    it("calls next() after successfully reading text", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "Content",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });

    it("does not set aborted on success", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "Content",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBeUndefined();
      expect(ctx.abortReason).toBeUndefined();
    });
  });

  // -----------------------------------------------------------------------
  // Happy path — selection text
  // -----------------------------------------------------------------------

  describe("happy path: selection text", () => {
    it("sets ctx.isSelection to true when text comes from a selection", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "Selected text",
        isSelection: true,
      });

      await handler.handle(ctx, next);

      expect(ctx.isSelection).toBe(true);
      expect(ctx.text).toBe("Selected text");
    });

    it("calls next() for selection text", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "Selected",
        isSelection: true,
      });

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Emitter events
  // -----------------------------------------------------------------------

  describe("emitter events", () => {
    it("emits phaseStart('reading') at the beginning", async () => {
      const emitter = new PipelineEventEmitter();
      const onPhaseStart = vi.fn();
      emitter.subscribe({ onPhaseStart });

      const ctx = makeContext({ emitter });
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "Content",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(onPhaseStart).toHaveBeenCalledWith("reading", "Leyendo texto...");
    });

    it("emits phaseComplete('reading') on success", async () => {
      const emitter = new PipelineEventEmitter();
      const onPhaseComplete = vi.fn();
      emitter.subscribe({ onPhaseComplete });

      const ctx = makeContext({ emitter });
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "Content",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(onPhaseComplete).toHaveBeenCalledWith("reading");
    });

    it("emits phaseStart but NOT phaseComplete when text is empty", async () => {
      const emitter = new PipelineEventEmitter();
      const onPhaseStart = vi.fn();
      const onPhaseComplete = vi.fn();
      emitter.subscribe({ onPhaseStart, onPhaseComplete });

      const ctx = makeContext({ emitter });
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(onPhaseStart).toHaveBeenCalledOnce();
      expect(onPhaseComplete).not.toHaveBeenCalled();
    });

    it("emits abort with reason when text is empty", async () => {
      const emitter = new PipelineEventEmitter();
      const onAbort = vi.fn();
      emitter.subscribe({ onAbort });

      const ctx = makeContext({ emitter });
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(onAbort).toHaveBeenCalledWith(
        "El documento está vacío. Escribe algo primero."
      );
    });
  });

  // -----------------------------------------------------------------------
  // Abort — empty text
  // -----------------------------------------------------------------------

  describe("abort: empty text", () => {
    it("aborts when text is an empty string", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBe(true);
      expect(ctx.abortReason).toBe(
        "El documento está vacío. Escribe algo primero."
      );
      expect(next).not.toHaveBeenCalled();
    });

    it("aborts when text is only whitespace", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "   \n\t  ",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBe(true);
      expect(next).not.toHaveBeenCalled();
    });

    it("does not set ctx.text when aborting on empty text", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      // The handler returns before setting ctx.text in the abort path
      expect(ctx.text).toBeUndefined();
    });
  });

  // -----------------------------------------------------------------------
  // Edge cases
  // -----------------------------------------------------------------------

  describe("edge cases", () => {
    it("handles text with only a single character", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "a",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(ctx.text).toBe("a");
      expect(ctx.aborted).toBeUndefined();
      expect(next).toHaveBeenCalledOnce();
    });

    it("propagates errors thrown by documentPort.getTextToAnalyze", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockRejectedValue(
        new Error("Office.js error")
      );

      await expect(handler.handle(ctx, next)).rejects.toThrow(
        "Office.js error"
      );
      expect(next).not.toHaveBeenCalled();
    });

    it("calls documentPort.getTextToAnalyze exactly once", async () => {
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: "Content",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(ctx.documentPort.getTextToAnalyze).toHaveBeenCalledOnce();
    });

    it("handles very large text without issue", async () => {
      const largeText = "x".repeat(500_000);
      const ctx = makeContext();
      vi.mocked(ctx.documentPort.getTextToAnalyze).mockResolvedValue({
        text: largeText,
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(ctx.text).toBe(largeText);
      expect(ctx.text!.length).toBe(500_000);
      expect(next).toHaveBeenCalledOnce();
    });
  });
});
