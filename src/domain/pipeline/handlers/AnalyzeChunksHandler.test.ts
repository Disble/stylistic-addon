import type { IAnalysisPort, IDocumentPort } from "../../ports";
import type { ChunkPollResult, Suggestion, TextChunk } from "../../types";
import { DEFAULT_AUTHOR_SLUG } from "../../../infrastructure/config";
import type { PipelineContext } from "../PipelineContext";
import { PipelineEventEmitter } from "../PipelineEvents";
import { AnalyzeChunksHandler } from "./AnalyzeChunksHandler";
import { beforeEach, describe, expect, it, vi } from "vitest";

/** Builds a canonical suggestion fixture for analysis tests. */
function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  const anchor = overrides.anchor ?? "texto original";
  return {
    id: "s1",
    context: overrides.context ?? `Contexto con ${anchor}.`,
    anchor,
    suggestedText: "texto mejorado",
    justification: "Mejora de estilo",
    category: "Redundancia",
    severity: "medium",
    type: "track-change",
    ...overrides,
  };
}

/** Builds a canonical chunk fixture for analysis tests. */
function makeChunk(overrides: Partial<TextChunk> = {}): TextChunk {
  return {
    text: "Este es un párrafo de ejemplo para analizar.",
    index: 0,
    total: 1,
    startOffset: 0,
    ...overrides,
  };
}

/** Builds a poll result fixture with sane defaults. */
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

/** Creates a mocked analysis port for one handler run. */
function makeMockAnalysisPort(): IAnalysisPort {
  return {
    checkConnection: vi.fn().mockResolvedValue(true),
    submitChunkAnalysis: vi
      .fn()
      .mockResolvedValue({ chunkIndex: 0, runId: "run-0" }),
    pollChunkAnalysis: vi.fn().mockResolvedValue(makePollResult()),
  };
}

/** Creates a mocked document port for the shared pipeline context. */
function makeMockDocumentPort(): IDocumentPort {
  return {
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
  };
}

/** Creates the pipeline context consumed by AnalyzeChunksHandler. */
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

describe("AnalyzeChunksHandler", () => {
  let handler: AnalyzeChunksHandler;
  let next: ReturnType<typeof vi.fn<() => Promise<void>>>;

  beforeEach(() => {
    handler = new AnalyzeChunksHandler(0);
    next = vi.fn<() => Promise<void>>().mockResolvedValue(undefined);
  });

  it("submits and polls each chunk, then collects all suggestions", async () => {
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

    expect(analysisPort.submitChunkAnalysis).toHaveBeenNthCalledWith(
      1,
      chunks[0],
      "general",
      DEFAULT_AUTHOR_SLUG,
    );
    expect(analysisPort.submitChunkAnalysis).toHaveBeenNthCalledWith(
      2,
      chunks[1],
      "general",
      DEFAULT_AUTHOR_SLUG,
    );
    expect(analysisPort.pollChunkAnalysis).toHaveBeenCalledWith(0, "run-0");
    expect(analysisPort.pollChunkAnalysis).toHaveBeenCalledWith(1, "run-1");
    expect(ctx.rawSuggestions?.map((suggestion) => suggestion.id)).toEqual([
      "s-0",
      "s-1",
    ]);
    expect(ctx.chunkErrors).toEqual([]);
    expect(next).toHaveBeenCalledOnce();
  });

  it("keeps successful suggestions when another chunk fails", async () => {
    const analysisPort = makeMockAnalysisPort();
    (analysisPort.submitChunkAnalysis as ReturnType<typeof vi.fn>)
      .mockResolvedValueOnce({ chunkIndex: 0, runId: "run-0" })
      .mockResolvedValueOnce({ chunkIndex: 1, runId: "run-1" });
    (analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>)
      .mockResolvedValueOnce(
        makePollResult({
          chunkIndex: 0,
          runId: "run-0",
          suggestions: [makeSuggestion({ id: "ok" })],
        }),
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
      chunks: [makeChunk({ index: 0, total: 2 }), makeChunk({ index: 1, total: 2 })],
    });
    await handler.handle(ctx, next);

    expect(ctx.rawSuggestions?.map((suggestion) => suggestion.id)).toEqual(["ok"]);
    expect(ctx.chunkErrors).toEqual(["Backend timeout"]);
    expect(next).toHaveBeenCalledOnce();
  });

  it("aborts with a chunk-error message when no suggestions survive analysis", async () => {
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
      chunks: [makeChunk({ index: 0, total: 2 }), makeChunk({ index: 1, total: 2 })],
    });
    await handler.handle(ctx, next);

    expect(ctx.aborted).toBe(true);
    expect(ctx.abortReason).toContain("2 fragmento(s)");
    expect(emitAbort).toHaveBeenCalledWith(ctx.abortReason);
    expect(next).not.toHaveBeenCalled();
  });

  it("aborts with a no-suggestions message when analysis succeeds but finds nothing", async () => {
    const analysisPort = makeMockAnalysisPort();
    (analysisPort.pollChunkAnalysis as ReturnType<typeof vi.fn>).mockResolvedValue(
      makePollResult({ suggestions: [] }),
    );
    const ctx = makePipelineContext({ analysisPort });
    await handler.handle(ctx, next);

    expect(ctx.aborted).toBe(true);
    expect(ctx.abortReason).toBe("No se encontraron sugerencias editoriales.");
    expect(ctx.chunkErrors).toEqual([]);
    expect(next).not.toHaveBeenCalled();
  });

  it("emits queue and poll progress for each chunk", async () => {
    const emitter = new PipelineEventEmitter();
    const emitProgress = vi.spyOn(emitter, "emitProgress");
    const ctx = makePipelineContext({
      emitter,
      chunks: [makeChunk({ index: 0, total: 1 })],
    });
    await handler.handle(ctx, next);

    expect(emitProgress).toHaveBeenCalledWith(0, 2, "Encolando fragmento 1 de 1...");
    expect(emitProgress).toHaveBeenCalledWith(
      1,
      2,
      "Consultando resultado del fragmento 1 de 1...",
    );
  });
});
