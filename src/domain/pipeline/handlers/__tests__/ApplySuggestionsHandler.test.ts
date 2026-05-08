import type { IAnalysisPort, IDocumentPort } from "../../../ports";
import type { ApplySuggestionsResult } from "../../../DocumentApplication.types";
import type { Suggestion } from "../../../suggestion/Suggestion.types";
import type { ProgressCallback } from "../../PipelineEvents.types";
import type { PipelineContext } from "../../PipelineContext";
import { PipelineEventEmitter } from "../../PipelineEvents";
import { ApplySuggestionsHandler } from "../ApplySuggestionsHandler";

// ---------------------------------------------------------------------------
// Factories
// ---------------------------------------------------------------------------

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

function makeInsertionResult(
  overrides: Partial<ApplySuggestionsResult> = {}
): ApplySuggestionsResult {
  return {
    successCount: 1,
    failedSuggestions: [],
    pendingAfter: {
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    },
    documentState: "pending-review",
    trackChangesActivatedForBatch: false,
    ...overrides,
  };
}

function makeMockDocumentPort(
  result: ApplySuggestionsResult = makeInsertionResult()
): IDocumentPort {
  return {
    getTextToAnalyze: vi.fn(),
    getDocumentUuid: vi.fn(),
    getAppliedOriginalTexts: vi.fn(),
    applySuggestions: vi.fn().mockResolvedValue(result),
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

function makeMockAnalysisPort(): IAnalysisPort {
  return {
    checkConnection: vi.fn(),
    submitChunkAnalysis: vi.fn(),
    pollChunkAnalysis: vi.fn(),
  };
}

function makePipelineContext(
  pendingSuggestions: Suggestion[],
  documentPort?: IDocumentPort,
  overrides: Partial<PipelineContext> = {}
): PipelineContext {
  return {
    documentPort: documentPort ?? makeMockDocumentPort(),
    analysisPort: makeMockAnalysisPort(),
    emitter: new PipelineEventEmitter(),
    genero: "general",
    maxChunkSize: 100_000,
    pendingSuggestions,
    chunkErrors: [],
    isSelection: false,
    ...overrides,
  } as PipelineContext;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("ApplySuggestionsHandler", () => {
  let handler: ApplySuggestionsHandler;
  let next: () => Promise<void>;

  beforeEach(() => {
    handler = new ApplySuggestionsHandler();
    next = vi.fn<() => Promise<void>>();
  });

  // -----------------------------------------------------------------------
  // Happy path
  // -----------------------------------------------------------------------

  describe("happy path", () => {
    it("should call documentPort.applySuggestions with pending suggestions", async () => {
      const suggestions = [makeSuggestion({ id: "s1" }), makeSuggestion({ id: "s2" })];
      const docPort = makeMockDocumentPort(makeInsertionResult({ successCount: 2 }));
      const ctx = makePipelineContext(suggestions, docPort);

      await handler.handle(ctx, next);

      expect(docPort.applySuggestions).toHaveBeenCalledWith(suggestions, expect.any(Function));
    });

    it("should store the InsertionResult in ctx.result", async () => {
      const result = makeInsertionResult({
        successCount: 3,
        failedSuggestions: [],
      });
      const docPort = makeMockDocumentPort(result);
      const suggestions = [makeSuggestion()];
      const ctx = makePipelineContext(suggestions, docPort);

      await handler.handle(ctx, next);

      expect(ctx.result).toEqual(result);
    });

    it("should call next() after applying", async () => {
      const ctx = makePipelineContext([makeSuggestion()]);

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Event emission
  // -----------------------------------------------------------------------

  describe("event emission", () => {
    it("should emit phaseComplete with 'applying'", async () => {
      const emitter = new PipelineEventEmitter();
      const emitPhaseComplete = vi.spyOn(emitter, "emitPhaseComplete");
      const ctx = makePipelineContext([makeSuggestion()], undefined, {
        emitter,
      });

      await handler.handle(ctx, next);

      expect(emitPhaseComplete).toHaveBeenCalledWith("applying");
    });

    it("should emit complete with suggestions, result, chunkErrors, and isSelection", async () => {
      const emitter = new PipelineEventEmitter();
      const emitComplete = vi.spyOn(emitter, "emitComplete");

      const suggestions = [makeSuggestion({ id: "s1" })];
      const result = makeInsertionResult({ successCount: 1 });
      const docPort = makeMockDocumentPort(result);
      const chunkErrors = ["chunk 2 failed"];

      const ctx = makePipelineContext(suggestions, docPort, {
        emitter,
        chunkErrors,
        isSelection: true,
      });

      await handler.handle(ctx, next);

      expect(emitComplete).toHaveBeenCalledWith(suggestions, result, chunkErrors, true);
    });

    it("should pass isSelection=false when analyzing full document", async () => {
      const emitter = new PipelineEventEmitter();
      const emitComplete = vi.spyOn(emitter, "emitComplete");

      const ctx = makePipelineContext([makeSuggestion()], undefined, {
        emitter,
        isSelection: false,
      });

      await handler.handle(ctx, next);

      expect(emitComplete).toHaveBeenCalledWith(
        expect.any(Array),
        expect.any(Object),
        expect.any(Array),
        false
      );
    });

    it("should use empty array for chunkErrors when undefined", async () => {
      const emitter = new PipelineEventEmitter();
      const emitComplete = vi.spyOn(emitter, "emitComplete");

      const ctx = makePipelineContext([makeSuggestion()], undefined, {
        emitter,
        chunkErrors: undefined,
      });

      await handler.handle(ctx, next);

      // ctx.chunkErrors ?? [] in the handler
      expect(emitComplete).toHaveBeenCalledWith(
        expect.any(Array),
        expect.any(Object),
        [],
        expect.any(Boolean)
      );
    });
  });

  // -----------------------------------------------------------------------
  // Progress bridge
  // -----------------------------------------------------------------------

  describe("progress bridge", () => {
    it("should pass a ProgressCallback that bridges to emitter.emitProgress", async () => {
      const emitter = new PipelineEventEmitter();
      const emitProgress = vi.spyOn(emitter, "emitProgress");

      const docPort = makeMockDocumentPort();
      // Capture the onProgress callback and invoke it
      (docPort.applySuggestions as ReturnType<typeof vi.fn>).mockImplementation(
        async (_suggestions: Suggestion[], onProgress?: ProgressCallback) => {
          onProgress?.("applying", 1, 3, "Aplicando sugerencia 1 de 3...");
          onProgress?.("applying", 2, 3, "Aplicando sugerencia 2 de 3...");
          onProgress?.("applying", 3, 3, "Aplicando sugerencia 3 de 3...");
          return makeInsertionResult({ successCount: 3 });
        }
      );

      const ctx = makePipelineContext(
        [makeSuggestion(), makeSuggestion({ id: "s2" }), makeSuggestion({ id: "s3" })],
        docPort,
        { emitter }
      );

      await handler.handle(ctx, next);

      expect(emitProgress).toHaveBeenCalledTimes(3);
      expect(emitProgress).toHaveBeenNthCalledWith(1, 1, 3, "Aplicando sugerencia 1 de 3...");
      expect(emitProgress).toHaveBeenNthCalledWith(2, 2, 3, "Aplicando sugerencia 2 de 3...");
      expect(emitProgress).toHaveBeenNthCalledWith(3, 3, 3, "Aplicando sugerencia 3 de 3...");
    });
  });

  // -----------------------------------------------------------------------
  // Partial apply (some suggestions fail)
  // -----------------------------------------------------------------------

  describe("partial apply", () => {
    it("should store result with failed suggestions", async () => {
      const failedSuggestion = makeSuggestion({
        id: "failed-1",
        anchor: "no encontrado",
        context: "Contexto con no encontrado.",
      });
      const result = makeInsertionResult({
        successCount: 2,
        failedSuggestions: [
          {
            suggestion: failedSuggestion,
            reason: "not-found",
            message: "Anchor no encontrado en el contexto",
          },
        ],
      });
      const docPort = makeMockDocumentPort(result);

      const suggestions = [
        makeSuggestion({ id: "s1" }),
        makeSuggestion({ id: "s2" }),
        failedSuggestion,
      ];
      const ctx = makePipelineContext(suggestions, docPort);

      await handler.handle(ctx, next);

      expect(ctx.result?.successCount).toBe(2);
      expect(ctx.result?.failedSuggestions).toHaveLength(1);
      expect(ctx.result?.failedSuggestions[0].suggestion.id).toBe("failed-1");
    });

    it("should still call next() and emit events even with failures", async () => {
      const emitter = new PipelineEventEmitter();
      const emitPhaseComplete = vi.spyOn(emitter, "emitPhaseComplete");
      const emitComplete = vi.spyOn(emitter, "emitComplete");

      const result = makeInsertionResult({
        successCount: 0,
        failedSuggestions: [
          {
            suggestion: makeSuggestion(),
            reason: "command-error",
            message: "insert failed",
          },
        ],
      });
      const docPort = makeMockDocumentPort(result);
      const ctx = makePipelineContext([makeSuggestion()], docPort, { emitter });

      await handler.handle(ctx, next);

      expect(emitPhaseComplete).toHaveBeenCalledWith("applying");
      expect(emitComplete).toHaveBeenCalledOnce();
      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // All suggestions fail
  // -----------------------------------------------------------------------

  describe("all suggestions fail", () => {
    it("should store result with all failed and successCount 0", async () => {
      const failed = [makeSuggestion({ id: "f1" }), makeSuggestion({ id: "f2" })];
      const result = makeInsertionResult({
        successCount: 0,
        failedSuggestions: failed.map((suggestion) => ({
          suggestion,
          reason: "not-found" as const,
          message: "Anchor no encontrado en el contexto",
        })),
      });
      const docPort = makeMockDocumentPort(result);
      const ctx = makePipelineContext(failed, docPort);

      await handler.handle(ctx, next);

      expect(ctx.result?.successCount).toBe(0);
      expect(ctx.result?.failedSuggestions).toHaveLength(2);
      // Still calls next() — handler does not abort on failures
      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Handler does NOT abort (no aborted flag set)
  // -----------------------------------------------------------------------

  describe("abort behavior", () => {
    it("should never set ctx.aborted", async () => {
      const ctx = makePipelineContext([makeSuggestion()]);

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBeUndefined();
    });
  });

  // -----------------------------------------------------------------------
  // Single suggestion
  // -----------------------------------------------------------------------

  describe("single suggestion", () => {
    it("should apply a single suggestion correctly", async () => {
      const suggestion = makeSuggestion({ id: "only-one" });
      const result = makeInsertionResult({ successCount: 1 });
      const docPort = makeMockDocumentPort(result);
      const ctx = makePipelineContext([suggestion], docPort);

      await handler.handle(ctx, next);

      expect(docPort.applySuggestions).toHaveBeenCalledWith([suggestion], expect.any(Function));
      expect(ctx.result).toEqual(result);
    });
  });

  // -----------------------------------------------------------------------
  // Emitter receives chunkErrors from context
  // -----------------------------------------------------------------------

  describe("chunkErrors propagation", () => {
    it("should pass existing chunkErrors to emitComplete", async () => {
      const emitter = new PipelineEventEmitter();
      const emitComplete = vi.spyOn(emitter, "emitComplete");

      const chunkErrors = ["chunk 1 timeout", "chunk 3 malformed response"];
      const ctx = makePipelineContext([makeSuggestion()], undefined, {
        emitter,
        chunkErrors,
      });

      await handler.handle(ctx, next);

      expect(emitComplete).toHaveBeenCalledWith(
        expect.any(Array),
        expect.any(Object),
        chunkErrors,
        expect.any(Boolean)
      );
    });
  });
});
