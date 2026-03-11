import { AnalyzeChunksHandler } from "./AnalyzeChunksHandler";
import { PipelineContext } from "../PipelineContext";
import { PipelineEventEmitter } from "../PipelineEvents";
import type { IDocumentPort, IAnalysisPort } from "../../ports";
import type { Suggestion, TextChunk, ChunkResult } from "../../types";

// ---------------------------------------------------------------------------
// Factories
// ---------------------------------------------------------------------------

function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  return {
    id: "s1",
    originalText: "texto original",
    suggestedText: "texto mejorado",
    justification: "Mejora de estilo",
    category: "Redundancia",
    severity: "medium",
    ...overrides,
  };
}

function makeChunk(overrides: Partial<TextChunk> = {}): TextChunk {
  return {
    text: "Este es un párrafo de ejemplo para analizar.",
    index: 0,
    total: 1,
    startOffset: 0,
    ...overrides,
  };
}

function makeChunkResult(overrides: Partial<ChunkResult> = {}): ChunkResult {
  return {
    chunkIndex: 0,
    suggestions: [makeSuggestion()],
    ...overrides,
  };
}

function makeMockAnalysisPort(): IAnalysisPort {
  return {
    checkConnection: vi.fn().mockResolvedValue(true),
    analyzeChunk: vi.fn().mockResolvedValue(makeChunkResult()),
  };
}

function makeMockDocumentPort(): IDocumentPort {
  return {
    getTextToAnalyze: vi.fn(),
    getAppliedOriginalTexts: vi.fn(),
    applySuggestions: vi.fn(),
    cleanupResolvedComments: vi.fn(),
    acceptSuggestion: vi.fn(),
    rejectSuggestion: vi.fn(),
  };
}

function makePipelineContext(overrides: Partial<PipelineContext> = {}): PipelineContext {
  return {
    documentPort: makeMockDocumentPort(),
    analysisPort: makeMockAnalysisPort(),
    emitter: new PipelineEventEmitter(),
    profile: "general",
    maxChunkSize: 100_000,
    chunks: [makeChunk()],
    isSelection: false,
    ...overrides,
  } as PipelineContext;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("AnalyzeChunksHandler", () => {
  let handler: AnalyzeChunksHandler;
  let next: () => Promise<void>;

  beforeEach(() => {
    handler = new AnalyzeChunksHandler();
    next = vi.fn<() => Promise<void>>();
  });

  // -----------------------------------------------------------------------
  // Happy path
  // -----------------------------------------------------------------------

  describe("happy path", () => {
    it("should call analysisPort.analyzeChunk for each chunk", async () => {
      const chunks = [
        makeChunk({ index: 0, total: 2, text: "Chunk uno." }),
        makeChunk({ index: 1, total: 2, text: "Chunk dos." }),
      ];
      const ctx = makePipelineContext({ chunks });

      await handler.handle(ctx, next);

      expect(ctx.analysisPort.analyzeChunk).toHaveBeenCalledTimes(2);
      expect(ctx.analysisPort.analyzeChunk).toHaveBeenNthCalledWith(
        1,
        chunks[0],
        "general",
        "es"
      );
      expect(ctx.analysisPort.analyzeChunk).toHaveBeenNthCalledWith(
        2,
        chunks[1],
        "general",
        "es"
      );
    });

    it("should collect all suggestions into ctx.rawSuggestions", async () => {
      const s1 = makeSuggestion({ id: "c0-1", originalText: "foo" });
      const s2 = makeSuggestion({ id: "c1-1", originalText: "bar" });

      const analysisPort = makeMockAnalysisPort();
      (analysisPort.analyzeChunk as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(makeChunkResult({ chunkIndex: 0, suggestions: [s1] }))
        .mockResolvedValueOnce(makeChunkResult({ chunkIndex: 1, suggestions: [s2] }));

      const ctx = makePipelineContext({
        analysisPort,
        chunks: [
          makeChunk({ index: 0, total: 2 }),
          makeChunk({ index: 1, total: 2 }),
        ],
      });

      await handler.handle(ctx, next);

      expect(ctx.rawSuggestions).toEqual([s1, s2]);
    });

    it("should call next() when suggestions are found", async () => {
      const ctx = makePipelineContext();

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });

    it("should set chunkErrors to empty array when no errors occur", async () => {
      const ctx = makePipelineContext();

      await handler.handle(ctx, next);

      expect(ctx.chunkErrors).toEqual([]);
    });

    it("should use the profile from ctx", async () => {
      const ctx = makePipelineContext({ profile: "formal" });

      await handler.handle(ctx, next);

      expect(ctx.analysisPort.analyzeChunk).toHaveBeenCalledWith(
        expect.any(Object),
        "formal",
        "es"
      );
    });
  });

  // -----------------------------------------------------------------------
  // Progress emission
  // -----------------------------------------------------------------------

  describe("progress emission", () => {
    it("should emit phaseStart and progress for each chunk", async () => {
      const emitter = new PipelineEventEmitter();
      const emitPhaseStart = vi.spyOn(emitter, "emitPhaseStart");
      const emitProgress = vi.spyOn(emitter, "emitProgress");

      const chunks = [
        makeChunk({ index: 0, total: 3 }),
        makeChunk({ index: 1, total: 3 }),
        makeChunk({ index: 2, total: 3 }),
      ];
      const ctx = makePipelineContext({ emitter, chunks });

      await handler.handle(ctx, next);

      expect(emitPhaseStart).toHaveBeenCalledTimes(3);
      expect(emitProgress).toHaveBeenCalledTimes(3);
      expect(emitProgress).toHaveBeenNthCalledWith(
        1,
        1,
        3,
        "Analizando fragmento 1 de 3..."
      );
      expect(emitProgress).toHaveBeenNthCalledWith(
        2,
        2,
        3,
        "Analizando fragmento 2 de 3..."
      );
      expect(emitProgress).toHaveBeenNthCalledWith(
        3,
        3,
        3,
        "Analizando fragmento 3 de 3..."
      );
    });
  });

  // -----------------------------------------------------------------------
  // Partial failures (some chunks error, some succeed)
  // -----------------------------------------------------------------------

  describe("partial failures", () => {
    it("should continue collecting suggestions when some chunks fail", async () => {
      const s1 = makeSuggestion({ id: "c0-1" });

      const analysisPort = makeMockAnalysisPort();
      (analysisPort.analyzeChunk as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(makeChunkResult({ chunkIndex: 0, suggestions: [s1] }))
        .mockResolvedValueOnce(
          makeChunkResult({ chunkIndex: 1, suggestions: [], error: "Backend timeout" })
        );

      const ctx = makePipelineContext({
        analysisPort,
        chunks: [
          makeChunk({ index: 0, total: 2 }),
          makeChunk({ index: 1, total: 2 }),
        ],
      });

      await handler.handle(ctx, next);

      expect(ctx.rawSuggestions).toEqual([s1]);
      expect(ctx.chunkErrors).toEqual(["Backend timeout"]);
      expect(next).toHaveBeenCalledOnce();
    });

    it("should record multiple chunk errors", async () => {
      const s1 = makeSuggestion({ id: "ok" });
      const analysisPort = makeMockAnalysisPort();
      (analysisPort.analyzeChunk as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(
          makeChunkResult({ chunkIndex: 0, suggestions: [], error: "Error 1" })
        )
        .mockResolvedValueOnce(
          makeChunkResult({ chunkIndex: 1, suggestions: [s1] })
        )
        .mockResolvedValueOnce(
          makeChunkResult({ chunkIndex: 2, suggestions: [], error: "Error 2" })
        );

      const ctx = makePipelineContext({
        analysisPort,
        chunks: [
          makeChunk({ index: 0, total: 3 }),
          makeChunk({ index: 1, total: 3 }),
          makeChunk({ index: 2, total: 3 }),
        ],
      });

      await handler.handle(ctx, next);

      expect(ctx.chunkErrors).toEqual(["Error 1", "Error 2"]);
      expect(ctx.rawSuggestions).toEqual([s1]);
      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Zero suggestions — abort
  // -----------------------------------------------------------------------

  describe("zero suggestions", () => {
    it("should abort with chunk-error message when all chunks fail", async () => {
      const analysisPort = makeMockAnalysisPort();
      (analysisPort.analyzeChunk as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(
          makeChunkResult({ chunkIndex: 0, suggestions: [], error: "fail 1" })
        )
        .mockResolvedValueOnce(
          makeChunkResult({ chunkIndex: 1, suggestions: [], error: "fail 2" })
        );

      const emitter = new PipelineEventEmitter();
      const emitAbort = vi.spyOn(emitter, "emitAbort");

      const ctx = makePipelineContext({
        analysisPort,
        emitter,
        chunks: [
          makeChunk({ index: 0, total: 2 }),
          makeChunk({ index: 1, total: 2 }),
        ],
      });

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBe(true);
      expect(ctx.abortReason).toContain("2 fragmento(s)");
      expect(emitAbort).toHaveBeenCalledWith(ctx.abortReason);
      expect(next).not.toHaveBeenCalled();
    });

    it("should abort with no-suggestions message when chunks succeed but return nothing", async () => {
      const analysisPort = makeMockAnalysisPort();
      (analysisPort.analyzeChunk as ReturnType<typeof vi.fn>).mockResolvedValue(
        makeChunkResult({ suggestions: [] })
      );

      const emitter = new PipelineEventEmitter();
      const emitAbort = vi.spyOn(emitter, "emitAbort");

      const ctx = makePipelineContext({
        analysisPort,
        emitter,
        chunks: [makeChunk()],
      });

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBe(true);
      expect(ctx.abortReason).toBe("No se encontraron sugerencias editoriales.");
      expect(ctx.chunkErrors).toEqual([]);
      expect(emitAbort).toHaveBeenCalledWith(ctx.abortReason);
      expect(next).not.toHaveBeenCalled();
    });

    it("should not call next() when aborting", async () => {
      const analysisPort = makeMockAnalysisPort();
      (analysisPort.analyzeChunk as ReturnType<typeof vi.fn>).mockResolvedValue(
        makeChunkResult({ suggestions: [] })
      );

      const ctx = makePipelineContext({ analysisPort });

      await handler.handle(ctx, next);

      expect(next).not.toHaveBeenCalled();
    });
  });

  // -----------------------------------------------------------------------
  // Single chunk
  // -----------------------------------------------------------------------

  describe("single chunk", () => {
    it("should handle a single chunk correctly", async () => {
      const suggestion = makeSuggestion({ id: "c0-0" });
      const analysisPort = makeMockAnalysisPort();
      (analysisPort.analyzeChunk as ReturnType<typeof vi.fn>).mockResolvedValue(
        makeChunkResult({ suggestions: [suggestion] })
      );

      const ctx = makePipelineContext({
        analysisPort,
        chunks: [makeChunk({ index: 0, total: 1 })],
      });

      await handler.handle(ctx, next);

      expect(ctx.rawSuggestions).toEqual([suggestion]);
      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Multiple suggestions per chunk
  // -----------------------------------------------------------------------

  describe("multiple suggestions per chunk", () => {
    it("should accumulate all suggestions from a single chunk", async () => {
      const suggestions = [
        makeSuggestion({ id: "c0-0", originalText: "uno" }),
        makeSuggestion({ id: "c0-1", originalText: "dos" }),
        makeSuggestion({ id: "c0-2", originalText: "tres" }),
      ];

      const analysisPort = makeMockAnalysisPort();
      (analysisPort.analyzeChunk as ReturnType<typeof vi.fn>).mockResolvedValue(
        makeChunkResult({ suggestions })
      );

      const ctx = makePipelineContext({ analysisPort });

      await handler.handle(ctx, next);

      expect(ctx.rawSuggestions).toHaveLength(3);
      expect(ctx.rawSuggestions).toEqual(suggestions);
    });
  });

  // -----------------------------------------------------------------------
  // isSelection affects emitter message
  // -----------------------------------------------------------------------

  describe("scope label", () => {
    it("should use 'selección' in emitter message when isSelection is true", async () => {
      const emitter = new PipelineEventEmitter();
      const emitPhaseStart = vi.spyOn(emitter, "emitPhaseStart");

      const ctx = makePipelineContext({ emitter, isSelection: true });

      await handler.handle(ctx, next);

      expect(emitPhaseStart).toHaveBeenCalledWith(
        "analyzing",
        expect.stringContaining("selección")
      );
    });

    it("should use 'documento' in emitter message when isSelection is false", async () => {
      const emitter = new PipelineEventEmitter();
      const emitPhaseStart = vi.spyOn(emitter, "emitPhaseStart");

      const ctx = makePipelineContext({ emitter, isSelection: false });

      await handler.handle(ctx, next);

      expect(emitPhaseStart).toHaveBeenCalledWith(
        "analyzing",
        expect.stringContaining("documento")
      );
    });
  });

  // -----------------------------------------------------------------------
  // Sequential processing
  // -----------------------------------------------------------------------

  describe("sequential processing", () => {
    it("should process chunks sequentially, not in parallel", async () => {
      const callOrder: number[] = [];
      const analysisPort = makeMockAnalysisPort();
      (analysisPort.analyzeChunk as ReturnType<typeof vi.fn>).mockImplementation(
        async (chunk: TextChunk) => {
          callOrder.push(chunk.index);
          return makeChunkResult({ chunkIndex: chunk.index, suggestions: [makeSuggestion()] });
        }
      );

      const ctx = makePipelineContext({
        analysisPort,
        chunks: [
          makeChunk({ index: 0, total: 3 }),
          makeChunk({ index: 1, total: 3 }),
          makeChunk({ index: 2, total: 3 }),
        ],
      });

      await handler.handle(ctx, next);

      expect(callOrder).toEqual([0, 1, 2]);
    });
  });
});
