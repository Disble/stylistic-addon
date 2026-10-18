import type { IAnalysisPort, IDocumentPort } from "../../../ports";
import type { TextChunk } from "../../../chunking/TextChunk.types";
import type { AnalysisSnapshot, ChunkAnalysisResult } from "../../../analysis/Analysis.types";
import { AnalysisRequestError } from "../../../analysis/AnalysisRequestError";
import type { Suggestion } from "../../../suggestion/Suggestion.types";
import type { PipelineContext } from "../../PipelineContext";
import { PipelineEventEmitter } from "../../PipelineEvents";
import { AnalyzeChunksHandler } from "../AnalyzeChunksHandler";
import { beforeEach, describe, expect, it, vi } from "vitest";

const DOCUMENT_UUID = "11111111-1111-4111-8111-111111111111";

/** Builds a canonical suggestion fixture for analysis tests. */
function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  const anchor = overrides.anchor ?? "texto original";
  return {
    id: "0123456789abcdef",
    context: overrides.context ?? `Contexto con ${anchor}.`,
    anchor,
    suggestedText: "texto mejorado",
    justification: "Mejora de estilo",
    category: "estilo",
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

/** Builds one chunk outcome inside an analysis snapshot. */
function makeChunkResult(overrides: Partial<ChunkAnalysisResult> = {}): ChunkAnalysisResult {
  return {
    chunkIndex: 0,
    status: "completed",
    suggestions: [makeSuggestion()],
    dropped: 0,
    ...overrides,
  };
}

/** Builds an analysis snapshot fixture with sane defaults. */
function makeSnapshot(overrides: Partial<AnalysisSnapshot> = {}): AnalysisSnapshot {
  return {
    analysisId: "a1",
    status: "completed",
    chunks: [makeChunkResult()],
    ...overrides,
  };
}

/** Creates a mocked analysis port whose analysis is running, then completed. */
function makeMockAnalysisPort(): IAnalysisPort {
  return {
    getReadiness: vi.fn().mockResolvedValue({ ready: true }),
    startAnalysis: vi.fn().mockResolvedValue(
      makeSnapshot({
        status: "running",
        chunks: [makeChunkResult({ status: "queued", suggestions: [] })],
      })
    ),
    getAnalysis: vi.fn().mockResolvedValue(makeSnapshot()),
    cancelAnalysis: vi.fn(),
  };
}

/** Creates a mocked document port for the shared pipeline context. */
function makeMockDocumentPort(): IDocumentPort {
  return {
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
}

/** Creates the pipeline context consumed by AnalyzeChunksHandler. */
function makePipelineContext(overrides: Partial<PipelineContext> = {}): PipelineContext {
  return {
    documentPort: makeMockDocumentPort(),
    analysisPort: makeMockAnalysisPort(),
    emitter: new PipelineEventEmitter(),
    genero: "general",
    maxChunkSize: 100_000,
    chunks: [makeChunk()],
    isSelection: false,
    documentUuid: DOCUMENT_UUID,
    ...overrides,
  };
}

describe("AnalyzeChunksHandler", () => {
  let handler: AnalyzeChunksHandler;
  let next: ReturnType<typeof vi.fn<() => Promise<void>>>;

  beforeEach(() => {
    handler = new AnalyzeChunksHandler(0);
    next = vi.fn<() => Promise<void>>().mockResolvedValue(undefined);
  });

  it("starts one analysis for every chunk, polls it, then collects suggestions in chunk order", async () => {
    const chunks = [
      makeChunk({ index: 0, total: 2, text: "Chunk uno." }),
      makeChunk({ index: 1, total: 2, text: "Chunk dos." }),
    ];
    const analysisPort = makeMockAnalysisPort();
    vi.mocked(analysisPort.getAnalysis).mockResolvedValueOnce(
      makeSnapshot({
        chunks: [
          makeChunkResult({ chunkIndex: 0, suggestions: [makeSuggestion({ id: "s-0" })] }),
          makeChunkResult({ chunkIndex: 1, suggestions: [makeSuggestion({ id: "s-1" })] }),
        ],
      })
    );

    const ctx = makePipelineContext({ chunks, analysisPort });
    await handler.handle(ctx, next);

    expect(analysisPort.startAnalysis).toHaveBeenCalledOnce();
    expect(analysisPort.startAnalysis).toHaveBeenCalledWith({
      documentUuid: DOCUMENT_UUID,
      genre: "general",
      chunks,
    });
    expect(analysisPort.getAnalysis).toHaveBeenCalledWith("a1");
    expect(ctx.rawSuggestions?.map((suggestion) => suggestion.id)).toEqual(["s-0", "s-1"]);
    expect(ctx.chunkErrors).toEqual([]);
    expect(ctx.activeAnalysisId).toBeUndefined();
    expect(next).toHaveBeenCalledOnce();
  });

  it("keeps polling while the analysis is running", async () => {
    const analysisPort = makeMockAnalysisPort();
    vi.mocked(analysisPort.getAnalysis)
      .mockResolvedValueOnce(
        makeSnapshot({
          status: "running",
          chunks: [makeChunkResult({ status: "running", suggestions: [] })],
        })
      )
      .mockResolvedValueOnce(makeSnapshot());

    const ctx = makePipelineContext({ analysisPort });
    await handler.handle(ctx, next);

    expect(analysisPort.getAnalysis).toHaveBeenCalledTimes(2);
    expect(next).toHaveBeenCalledOnce();
  });

  it("keeps successful suggestions of a partial analysis and reports each failed chunk's real cause", async () => {
    const analysisPort = makeMockAnalysisPort();
    vi.mocked(analysisPort.getAnalysis).mockResolvedValueOnce(
      makeSnapshot({
        status: "partial",
        chunks: [
          makeChunkResult({ chunkIndex: 0, suggestions: [makeSuggestion({ id: "ok" })] }),
          makeChunkResult({
            chunkIndex: 1,
            status: "failed",
            suggestions: [],
            error: "The model request failed: 400 MissingSessionID",
          }),
        ],
      })
    );

    const ctx = makePipelineContext({
      analysisPort,
      chunks: [makeChunk({ index: 0, total: 2 }), makeChunk({ index: 1, total: 2 })],
    });
    await handler.handle(ctx, next);

    expect(ctx.rawSuggestions?.map((suggestion) => suggestion.id)).toEqual(["ok"]);
    expect(ctx.chunkErrors).toEqual([
      "Fragmento 2: The model request failed: 400 MissingSessionID",
    ]);
    expect(next).toHaveBeenCalledOnce();
  });

  it("aborts with the backend's real cause when the analysis fails", async () => {
    const analysisPort = makeMockAnalysisPort();
    vi.mocked(analysisPort.getAnalysis).mockResolvedValueOnce(
      makeSnapshot({
        status: "failed",
        chunks: [
          makeChunkResult({
            status: "failed",
            suggestions: [],
            error: "The model request failed: 400 MissingSessionID",
          }),
        ],
      })
    );

    const emitter = new PipelineEventEmitter();
    const emitAbort = vi.spyOn(emitter, "emitAbort");
    const ctx = makePipelineContext({ analysisPort, emitter });
    await handler.handle(ctx, next);

    expect(ctx.aborted).toBe(true);
    expect(ctx.abortReason).toBe(
      "El análisis falló: The model request failed: 400 MissingSessionID"
    );
    expect(ctx.chunkErrors).toEqual(["The model request failed: 400 MissingSessionID"]);
    expect(emitAbort).toHaveBeenCalledWith(ctx.abortReason);
    expect(next).not.toHaveBeenCalled();
  });

  it("aborts with a no-suggestions message when analysis succeeds but finds nothing", async () => {
    const analysisPort = makeMockAnalysisPort();
    vi.mocked(analysisPort.getAnalysis).mockResolvedValueOnce(
      makeSnapshot({ chunks: [makeChunkResult({ suggestions: [] })] })
    );
    const ctx = makePipelineContext({ analysisPort });
    await handler.handle(ctx, next);

    expect(ctx.aborted).toBe(true);
    expect(ctx.abortReason).toBe("No se encontraron sugerencias editoriales.");
    expect(ctx.chunkErrors).toEqual([]);
    expect(next).not.toHaveBeenCalled();
  });

  it("aborts with the backend's explanation when the analysis cannot start", async () => {
    const analysisPort = makeMockAnalysisPort();
    vi.mocked(analysisPort.startAnalysis).mockRejectedValueOnce(
      new AnalysisRequestError("No model is chosen in the hub.", 409)
    );
    const ctx = makePipelineContext({ analysisPort });
    await handler.handle(ctx, next);

    expect(ctx.aborted).toBe(true);
    expect(ctx.abortReason).toBe("No se pudo iniciar el análisis: No model is chosen in the hub.");
    expect(ctx.chunkErrors).toEqual(["No model is chosen in the hub."]);
    expect(analysisPort.getAnalysis).not.toHaveBeenCalled();
    expect(next).not.toHaveBeenCalled();
  });

  it("keeps the analysis for a retry query when polling fails locally", async () => {
    const analysisPort = makeMockAnalysisPort();
    vi.mocked(analysisPort.getAnalysis).mockRejectedValueOnce(
      new AnalysisRequestError("Failed to fetch")
    );

    const ctx = makePipelineContext({ analysisPort });
    await handler.handle(ctx, next);

    expect(ctx.retryableAnalysisId).toBe("a1");
    expect(ctx.activeAnalysisId).toBeUndefined();
    expect(ctx.aborted).toBe(true);
    expect(ctx.abortReason).toBe(
      "La consulta del análisis falló localmente: Failed to fetch. Reintentá la consulta."
    );
    expect(next).not.toHaveBeenCalled();
  });

  it("stops polling when the user cancels", async () => {
    const analysisPort = makeMockAnalysisPort();
    const ctx = makePipelineContext({ analysisPort, shouldCancelAnalysis: () => true });

    await handler.handle(ctx, next);

    expect(analysisPort.getAnalysis).not.toHaveBeenCalled();
    expect(ctx.aborted).toBe(true);
    expect(ctx.abortReason).toBe("Análisis cancelado por el usuario.");
    expect(ctx.activeAnalysisId).toBeUndefined();
    expect(next).not.toHaveBeenCalled();
  });

  it("aborts as canceled when the backend reports the analysis canceled", async () => {
    const analysisPort = makeMockAnalysisPort();
    vi.mocked(analysisPort.getAnalysis).mockResolvedValueOnce(
      makeSnapshot({
        status: "canceled",
        chunks: [makeChunkResult({ status: "canceled", suggestions: [] })],
      })
    );
    const ctx = makePipelineContext({ analysisPort });

    await handler.handle(ctx, next);

    expect(ctx.aborted).toBe(true);
    expect(ctx.abortReason).toBe("Análisis cancelado por el usuario.");
    expect(ctx.chunkErrors).toEqual([]);
  });

  it("publishes the active analysis before polling progress is observed, with finished-chunk counts", async () => {
    const emitter = new PipelineEventEmitter();
    const observed: Array<[number, number, string, string | undefined]> = [];
    const ctx = makePipelineContext({ emitter });
    emitter.subscribe({
      onPhaseStart: vi.fn(),
      onAbort: vi.fn(),
      onComplete: vi.fn(),
      onProgress: (current, total, message) => {
        observed.push([current, total, message, ctx.activeAnalysisId]);
      },
    });

    await handler.handle(ctx, next);

    expect(observed).toEqual([
      [0, 1, "Analizando 1 fragmento(s)...", "a1"],
      [1, 1, "Fragmentos terminados: 1 de 1.", "a1"],
    ]);
  });
});
