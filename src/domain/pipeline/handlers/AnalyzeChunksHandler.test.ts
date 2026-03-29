import type { IAnalysisPort, IDocumentPort } from "../../ports";
import type { ChunkPollResult, Suggestion, TextChunk } from "../../types";
import type { PipelineContext } from "../PipelineContext";
import { PipelineEventEmitter } from "../PipelineEvents";
import { AnalyzeChunksHandler } from "./AnalyzeChunksHandler";

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
    type: "track-change",
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

function makePollResult(
  overrides: Partial<ChunkPollResult> = {},
): ChunkPollResult {
  return {
    chunkIndex: 0,
    runId: "run-0",
    status: "success",
    suggestions: [makeSuggestion()],
    ...overrides,
  };
}

function makeMockAnalysisPort(): IAnalysisPort {
  return {
    checkConnection: vi.fn().mockResolvedValue(true),
    submitChunkAnalysis: vi
      .fn()
      .mockResolvedValue({ chunkIndex: 0, runId: "run-0" }),
    pollChunkAnalysis: vi.fn().mockResolvedValue(makePollResult()),
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

function makePipelineContext(
  overrides: Partial<PipelineContext> = {},
): PipelineContext {
  return {
    documentPort: makeMockDocumentPort(),
    analysisPort: makeMockAnalysisPort(),
    emitter: new PipelineEventEmitter(),
    genero: "general",
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
    handler = new AnalyzeChunksHandler(0);
    next = vi.fn<() => Promise<void>>();
  });

  // -----------------------------------------------------------------------
  // Happy path
  // -----------------------------------------------------------------------

  describe("happy path", () => {
    it("should submit each chunk and poll each run to completion", async () => {
      const chunks = [
        makeChunk({ index: 0, total: 2, text: "Chunk uno." }),
        makeChunk({ index: 1, total: 2, text: "Chunk dos." }),
      ];
      const analysisPort = makeMockAnalysisPort();
      (analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce({ chunkIndex: 0, runId: "run-0" })
        .mockResolvedValueOnce({ chunkIndex: 1, runId: "run-1" });
      (analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 0,
            runId: "run-0",
            suggestions: [makeSuggestion({ id: "s-0" })],
          }),
        )
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 1,
            runId: "run-1",
            suggestions: [makeSuggestion({ id: "s-1" })],
          }),
        );
      const ctx = makePipelineContext({ chunks, analysisPort });

      await handler.handle(ctx, next);

      expect(ctx.analysisPort.submitChunkAnalysis).toHaveBeenCalledTimes(2);
      expect(ctx.analysisPort.submitChunkAnalysis).toHaveBeenNthCalledWith(
        1,
        chunks[0],
        "general",
        "Disble",
      );
      expect(ctx.analysisPort.submitChunkAnalysis).toHaveBeenNthCalledWith(
        2,
        chunks[1],
        "general",
        "Disble",
      );
      expect(ctx.analysisPort.pollChunkAnalysis).toHaveBeenCalledWith(
        0,
        "run-0",
      );
      expect(ctx.analysisPort.pollChunkAnalysis).toHaveBeenCalledWith(
        1,
        "run-1",
      );
    });

    it("should collect all suggestions into ctx.rawSuggestions", async () => {
      const s1 = makeSuggestion({ id: "c0-1", originalText: "foo" });
      const s2 = makeSuggestion({ id: "c1-1", originalText: "bar" });

      const analysisPort = makeMockAnalysisPort();
      (analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce({ chunkIndex: 0, runId: "run-0" })
        .mockResolvedValueOnce({ chunkIndex: 1, runId: "run-1" });
      (analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(
          makePollResult({ chunkIndex: 0, runId: "run-0", suggestions: [s1] }),
        )
        .mockResolvedValueOnce(
          makePollResult({ chunkIndex: 1, runId: "run-1", suggestions: [s2] }),
        );

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

    it("should use the genero from ctx", async () => {
      const ctx = makePipelineContext({ genero: "narrativa-literaria" });

      await handler.handle(ctx, next);

      expect(ctx.analysisPort.submitChunkAnalysis).toHaveBeenCalledWith(
        expect.any(Object),
        "narrativa-literaria",
        "Disble",
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
      expect(emitProgress).toHaveBeenCalled();
      expect(emitProgress).toHaveBeenCalledWith(
        0,
        6,
        "Encolando fragmento 1 de 3...",
      );
      expect(emitProgress).toHaveBeenCalledWith(
        3,
        6,
        "Consultando resultado del fragmento 1 de 3...",
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
      (analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce({ chunkIndex: 0, runId: "run-0" })
        .mockResolvedValueOnce({ chunkIndex: 1, runId: "run-1" });
      (analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(
          makePollResult({ chunkIndex: 0, runId: "run-0", suggestions: [s1] }),
        )
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 1,
            runId: "run-1",
            status: "failed",
            suggestions: [],
            error: "Backend timeout",
          }),
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
      (analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce({ chunkIndex: 0, runId: "run-0" })
        .mockResolvedValueOnce({ chunkIndex: 1, runId: "run-1" })
        .mockResolvedValueOnce({ chunkIndex: 2, runId: "run-2" });
      (analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 0,
            runId: "run-0",
            status: "failed",
            suggestions: [],
            error: "Error 1",
          }),
        )
        .mockResolvedValueOnce(
          makePollResult({ chunkIndex: 1, runId: "run-1", suggestions: [s1] }),
        )
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 2,
            runId: "run-2",
            status: "failed",
            suggestions: [],
            error: "Error 2",
          }),
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

    it("should stop polling suspended runs and record them as chunk errors", async () => {
      const analysisPort = makeMockAnalysisPort();
      (
        analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockResolvedValue({
        chunkIndex: 0,
        runId: "run-0",
      });
      (
        analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockResolvedValue(
        makePollResult({
          chunkIndex: 0,
          runId: "run-0",
          status: "failed",
          suggestions: [],
          error:
            'Workflow entered "suspended" state and requires resume(), which this frontend does not support',
        }),
      );

      const ctx = makePipelineContext({
        analysisPort,
        chunks: [makeChunk()],
      });

      await handler.handle(ctx, next);

      expect(analysisPort.pollChunkAnalysis).toHaveBeenCalledTimes(1);
      expect(ctx.chunkErrors).toEqual([
        'Workflow entered "suspended" state and requires resume(), which this frontend does not support',
      ]);
      expect(ctx.aborted).toBe(true);
    });
  });

  // -----------------------------------------------------------------------
  // Zero suggestions — abort
  // -----------------------------------------------------------------------

  describe("zero suggestions", () => {
    it("should abort with chunk-error message when all chunks fail", async () => {
      const analysisPort = makeMockAnalysisPort();
      (analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce({ chunkIndex: 0, runId: "run-0" })
        .mockResolvedValueOnce({ chunkIndex: 1, runId: "run-1" });
      (analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 0,
            runId: "run-0",
            status: "failed",
            suggestions: [],
            error: "fail 1",
          }),
        )
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 1,
            runId: "run-1",
            status: "failed",
            suggestions: [],
            error: "fail 2",
          }),
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
      (
        analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockResolvedValue({
        chunkIndex: 0,
        runId: "run-0",
      });
      (
        analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockResolvedValue(makePollResult({ suggestions: [] }));

      const emitter = new PipelineEventEmitter();
      const emitAbort = vi.spyOn(emitter, "emitAbort");

      const ctx = makePipelineContext({
        analysisPort,
        emitter,
        chunks: [makeChunk()],
      });

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBe(true);
      expect(ctx.abortReason).toBe(
        "No se encontraron sugerencias editoriales.",
      );
      expect(ctx.chunkErrors).toEqual([]);
      expect(emitAbort).toHaveBeenCalledWith(ctx.abortReason);
      expect(next).not.toHaveBeenCalled();
    });

    it("should not call next() when aborting", async () => {
      const analysisPort = makeMockAnalysisPort();
      (
        analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockResolvedValue({
        chunkIndex: 0,
        runId: "run-0",
      });
      (
        analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockResolvedValue(makePollResult({ suggestions: [] }));

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
      (
        analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockResolvedValue({
        chunkIndex: 0,
        runId: "run-0",
      });
      (
        analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockResolvedValue(makePollResult({ suggestions: [suggestion] }));

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
      (
        analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockResolvedValue({
        chunkIndex: 0,
        runId: "run-0",
      });
      (
        analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockResolvedValue(makePollResult({ suggestions }));

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
        expect.stringContaining("selección"),
      );
    });

    it("should use 'documento' in emitter message when isSelection is false", async () => {
      const emitter = new PipelineEventEmitter();
      const emitPhaseStart = vi.spyOn(emitter, "emitPhaseStart");

      const ctx = makePipelineContext({ emitter, isSelection: false });

      await handler.handle(ctx, next);

      expect(emitPhaseStart).toHaveBeenCalledWith(
        "analyzing",
        expect.stringContaining("documento"),
      );
    });
  });

  // -----------------------------------------------------------------------
  // Sequential processing
  // -----------------------------------------------------------------------

  describe("sequential processing", () => {
    it("should submit chunks sequentially and poll them round-robin", async () => {
      const callOrder: number[] = [];
      const analysisPort = makeMockAnalysisPort();
      (
        analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>
      ).mockImplementation(async (chunk: TextChunk) => {
        callOrder.push(chunk.index);
        return { chunkIndex: chunk.index, runId: `run-${chunk.index}` };
      });
      (analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 0,
            runId: "run-0",
            status: "running",
            suggestions: [],
          }),
        )
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 1,
            runId: "run-1",
            suggestions: [makeSuggestion({ id: "s-1" })],
          }),
        )
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 2,
            runId: "run-2",
            suggestions: [makeSuggestion({ id: "s-2" })],
          }),
        )
        .mockResolvedValueOnce(
          makePollResult({
            chunkIndex: 0,
            runId: "run-0",
            suggestions: [makeSuggestion({ id: "s-0" })],
          }),
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
      expect(
        (analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>).mock.calls,
      ).toEqual([
        [0, "run-0"],
        [1, "run-1"],
        [2, "run-2"],
        [0, "run-0"],
      ]);
    });
  });
});
