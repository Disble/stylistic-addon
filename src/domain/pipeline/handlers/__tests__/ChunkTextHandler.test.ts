import type { IAnalysisPort, IDocumentPort } from "../../../ports";
import type { PipelineContext } from "../../PipelineContext";
import { PipelineEventEmitter } from "../../PipelineEvents";
import { ChunkTextHandler } from "../ChunkTextHandler";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeContext(overrides: Partial<PipelineContext> = {}): PipelineContext {
  const documentPort: IDocumentPort = {
    getTextToAnalyze: vi.fn(),
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
    maxChunkSize: 5000,
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("ChunkTextHandler", () => {
  let handler: ChunkTextHandler;
  let next: () => Promise<void>;

  beforeEach(() => {
    handler = new ChunkTextHandler();
    next = vi.fn<() => Promise<void>>();
  });

  // -----------------------------------------------------------------------
  // Happy path — single chunk
  // -----------------------------------------------------------------------

  describe("happy path: single chunk (text fits within maxChunkSize)", () => {
    it("produces a single chunk for short text", async () => {
      const ctx = makeContext({
        text: "Hello world",
        isSelection: false,
        maxChunkSize: 5000,
      });

      await handler.handle(ctx, next);

      expect(ctx.chunks).toHaveLength(1);
      expect(ctx.chunks?.[0].text).toBe("Hello world");
    });

    it("sets correct metadata on the single chunk", async () => {
      const ctx = makeContext({
        text: "Some paragraph",
        isSelection: false,
        maxChunkSize: 5000,
      });

      await handler.handle(ctx, next);

      const chunk = ctx.chunks?.[0];
      expect(chunk.index).toBe(0);
      expect(chunk.total).toBe(1);
      expect(chunk.startOffset).toBe(0);
    });

    it("calls next() after chunking", async () => {
      const ctx = makeContext({
        text: "Content",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Multiple chunks — text exceeds maxChunkSize
  // -----------------------------------------------------------------------

  describe("multiple chunks: text exceeds maxChunkSize", () => {
    it("splits text into multiple chunks at paragraph boundaries", async () => {
      const paragraph1 = "First paragraph.";
      const paragraph2 = "Second paragraph.";
      const text = `${paragraph1}\n\n${paragraph2}`;

      const ctx = makeContext({
        text,
        isSelection: false,
        // maxChunkSize just big enough for one paragraph but not both
        maxChunkSize: 20,
      });

      await handler.handle(ctx, next);

      expect(ctx.chunks?.length).toBeGreaterThanOrEqual(2);
      expect(ctx.chunks?.[0].text).toBe(paragraph1);
      expect(ctx.chunks?.[1].text).toBe(paragraph2);
    });

    it("sets correct total on all chunks", async () => {
      const text = "Paragraph A.\n\nParagraph B.\n\nParagraph C.";

      const ctx = makeContext({
        text,
        isSelection: false,
        maxChunkSize: 15,
      });

      await handler.handle(ctx, next);

      const total = ctx.chunks?.length;
      for (const chunk of ctx.chunks!) {
        expect(chunk.total).toBe(total);
      }
    });

    it("sets sequential index values starting from 0", async () => {
      const text = "AAA.\n\nBBB.\n\nCCC.";

      const ctx = makeContext({
        text,
        isSelection: false,
        maxChunkSize: 5,
      });

      await handler.handle(ctx, next);

      for (let i = 0; i < ctx.chunks?.length; i++) {
        expect(ctx.chunks?.[i].index).toBe(i);
      }
    });

    it("still calls next() with multiple chunks", async () => {
      const ctx = makeContext({
        text: "A.\n\nB.\n\nC.",
        isSelection: false,
        maxChunkSize: 3,
      });

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Uses ctx.maxChunkSize
  // -----------------------------------------------------------------------

  describe("respects ctx.maxChunkSize", () => {
    it("uses the maxChunkSize from context, not a hardcoded default", async () => {
      const paragraph = "x".repeat(100);
      const text = `${paragraph}\n\n${paragraph}`;

      // Large maxChunkSize → everything fits in one chunk
      const ctxLarge = makeContext({
        text,
        isSelection: false,
        maxChunkSize: 10_000,
      });
      await handler.handle(ctxLarge, vi.fn<() => Promise<void>>());
      expect(ctxLarge.chunks).toHaveLength(1);

      // Small maxChunkSize → splits into multiple chunks
      const ctxSmall = makeContext({
        text,
        isSelection: false,
        maxChunkSize: 150,
      });
      await handler.handle(ctxSmall, vi.fn<() => Promise<void>>());
      expect(ctxSmall.chunks?.length).toBeGreaterThan(1);
    });
  });

  // -----------------------------------------------------------------------
  // Edge cases
  // -----------------------------------------------------------------------

  describe("edge cases", () => {
    it("handles text with a single paragraph (no double newlines)", async () => {
      const ctx = makeContext({
        text: "One continuous paragraph without breaks.",
        isSelection: false,
        maxChunkSize: 5000,
      });

      await handler.handle(ctx, next);

      expect(ctx.chunks).toHaveLength(1);
      expect(ctx.chunks?.[0].text).toBe("One continuous paragraph without breaks.");
    });

    it("handles text from a selection (isSelection = true)", async () => {
      const ctx = makeContext({
        text: "Selected text.",
        isSelection: true,
        maxChunkSize: 5000,
      });

      await handler.handle(ctx, next);

      expect(ctx.chunks).toHaveLength(1);
      expect(next).toHaveBeenCalledOnce();
    });

    it("handles a single paragraph that exceeds maxChunkSize", async () => {
      // A paragraph bigger than maxChunkSize is sent as-is (per chunker design)
      const bigParagraph = "x".repeat(200);
      const ctx = makeContext({
        text: bigParagraph,
        isSelection: false,
        maxChunkSize: 50,
      });

      await handler.handle(ctx, next);

      // Should produce exactly 1 chunk because there's no paragraph boundary
      expect(ctx.chunks).toHaveLength(1);
      expect(ctx.chunks?.[0].text).toBe(bigParagraph);
    });
  });

  // -----------------------------------------------------------------------
  // Does not emit events or abort
  // -----------------------------------------------------------------------

  describe("no events or abort", () => {
    it("does not set ctx.aborted", async () => {
      const ctx = makeContext({
        text: "Content",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBeUndefined();
    });

    it("does not emit any phase events", async () => {
      const emitter = new PipelineEventEmitter();
      const onPhaseStart = vi.fn();
      const onPhaseComplete = vi.fn();
      const onAbort = vi.fn();
      emitter.subscribe({ onPhaseStart, onPhaseComplete, onAbort });

      const ctx = makeContext({
        text: "Content",
        isSelection: false,
        emitter,
      });

      await handler.handle(ctx, next);

      expect(onPhaseStart).not.toHaveBeenCalled();
      expect(onPhaseComplete).not.toHaveBeenCalled();
      expect(onAbort).not.toHaveBeenCalled();
    });
  });

  // -----------------------------------------------------------------------
  // Does not call any ports
  // -----------------------------------------------------------------------

  describe("isolation", () => {
    it("does not call any documentPort or analysisPort methods", async () => {
      const ctx = makeContext({
        text: "Content",
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(ctx.documentPort.getTextToAnalyze).not.toHaveBeenCalled();
      expect(ctx.documentPort.getAppliedOriginalTexts).not.toHaveBeenCalled();
      expect(ctx.documentPort.applySuggestions).not.toHaveBeenCalled();
      expect(ctx.documentPort.cleanupResolvedComments).not.toHaveBeenCalled();
      expect(ctx.analysisPort.checkConnection).not.toHaveBeenCalled();
      expect(ctx.analysisPort.submitChunkAnalysis).not.toHaveBeenCalled();
      expect(ctx.analysisPort.pollChunkAnalysis).not.toHaveBeenCalled();
    });
  });
});
