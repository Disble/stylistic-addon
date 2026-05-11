import type { IAnalysisPort, IDocumentPort } from "../../../ports";
import type { Suggestion } from "../../../suggestion/Suggestion.types";
import type { PipelineContext } from "../../PipelineContext";
import { PipelineEventEmitter } from "../../PipelineEvents";
import { DeduplicateHandler } from "../DeduplicateHandler";

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

function makeCommentOnlySuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  const anchor = overrides.anchor ?? "texto original";
  return {
    id: "c1",
    context: overrides.context ?? `Contexto con ${anchor}.`,
    anchor,
    justification: "Observación de estilo",
    category: "Registro",
    severity: "low",
    type: "comment-only",
    ...overrides,
  };
}

function makeMockPorts(): {
  documentPort: IDocumentPort;
  analysisPort: IAnalysisPort;
} {
  return {
    documentPort: {
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
    },
    analysisPort: {
      checkConnection: vi.fn(),
      submitChunkAnalysis: vi.fn(),
      pollChunkAnalysis: vi.fn(),
      cancelChunkAnalysis: vi.fn(),
      retryPollChunkAnalysis: vi.fn(),
    },
  };
}

function makePipelineContext(
  rawSuggestions: Suggestion[],
  overrides: Partial<PipelineContext> = {}
): PipelineContext {
  const { documentPort, analysisPort } = makeMockPorts();
  return {
    documentPort,
    analysisPort,
    emitter: new PipelineEventEmitter(),
    genero: "general",
    maxChunkSize: 100_000,
    rawSuggestions,
    ...overrides,
  } as PipelineContext;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("DeduplicateHandler", () => {
  let handler: DeduplicateHandler;
  let next: () => Promise<void>;

  beforeEach(() => {
    handler = new DeduplicateHandler();
    next = vi.fn<() => Promise<void>>();
  });

  // -----------------------------------------------------------------------
  // Happy path — no duplicates
  // -----------------------------------------------------------------------

  describe("no duplicates", () => {
    it("should pass through all suggestions when none are duplicated", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", anchor: "alfa", context: "Contexto alfa" }),
        makeSuggestion({ id: "s2", anchor: "beta", context: "Contexto beta" }),
        makeSuggestion({ id: "s3", anchor: "gamma", context: "Contexto gamma" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(3);
      expect(ctx.uniqueSuggestions).toEqual(suggestions);
    });

    it("should always call next()", async () => {
      const ctx = makePipelineContext([makeSuggestion()]);

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Duplicate detection
  // -----------------------------------------------------------------------

  describe("duplicate detection", () => {
    it("should remove exact duplicate context+anchor pairs", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", anchor: "repetido", context: "Contexto repetido" }),
        makeSuggestion({ id: "s2", anchor: "repetido", context: "Contexto repetido" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(1);
      expect(ctx.uniqueSuggestions?.[0].id).toBe("s1");
    });

    it("should keep the first occurrence and discard subsequent duplicates", async () => {
      const suggestions = [
        makeSuggestion({
          id: "first",
          anchor: "duplicado",
          context: "Contexto duplicado",
          suggestedText: "versión 1",
        }),
        makeSuggestion({
          id: "second",
          anchor: "duplicado",
          context: "Contexto duplicado",
          suggestedText: "versión 2",
        }),
        makeSuggestion({
          id: "third",
          anchor: "duplicado",
          context: "Contexto duplicado",
          suggestedText: "versión 3",
        }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(1);
      expect(ctx.uniqueSuggestions?.[0].id).toBe("first");
      expect(ctx.uniqueSuggestions?.[0].suggestedText).toBe("versión 1");
    });

    it("should handle a mix of unique and duplicate suggestions", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", anchor: "alfa", context: "Contexto alfa" }),
        makeSuggestion({ id: "s2", anchor: "beta", context: "Contexto beta" }),
        makeSuggestion({ id: "s3", anchor: "alfa", context: "Contexto alfa" }), // dup of s1
        makeSuggestion({ id: "s4", anchor: "gamma", context: "Contexto gamma" }),
        makeSuggestion({ id: "s5", anchor: "beta", context: "Contexto beta" }), // dup of s2
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(3);
      expect(ctx.uniqueSuggestions?.map((s) => s.id)).toEqual(["s1", "s2", "s4"]);
    });
  });

  // -----------------------------------------------------------------------
  // Case-insensitive comparison
  // -----------------------------------------------------------------------

  describe("case-insensitive deduplication", () => {
    it("should treat different cases as duplicates", async () => {
      const suggestions = [
        makeSuggestion({ id: "lower", anchor: "texto", context: "Contexto texto" }),
        makeSuggestion({ id: "upper", anchor: "TEXTO", context: "Contexto TEXTO" }),
        makeSuggestion({ id: "mixed", anchor: "Texto", context: "Contexto Texto" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(1);
      expect(ctx.uniqueSuggestions?.[0].id).toBe("lower");
    });

    it("should preserve the original casing of the kept suggestion", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", anchor: "MiTexto", context: "Contexto MiTexto" }),
        makeSuggestion({ id: "s2", anchor: "mitexto", context: "Contexto mitexto" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions?.[0].anchor).toBe("MiTexto");
    });
  });

  it("keeps same anchor when context differs", async () => {
    const suggestions = [
      makeSuggestion({ id: "s1", anchor: "texto", context: "Primer contexto con texto" }),
      makeSuggestion({ id: "s2", anchor: "texto", context: "Segundo contexto con texto" }),
    ];
    const ctx = makePipelineContext(suggestions);

    await handler.handle(ctx, next);

    expect(ctx.uniqueSuggestions).toHaveLength(2);
    expect(ctx.uniqueSuggestions?.map((s) => s.id)).toEqual(["s1", "s2"]);
  });

  // -----------------------------------------------------------------------
  // All duplicates
  // -----------------------------------------------------------------------

  describe("all duplicates", () => {
    it("should keep exactly one when all suggestions are duplicates", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", anchor: "mismo", context: "Contexto mismo" }),
        makeSuggestion({ id: "s2", anchor: "mismo", context: "Contexto mismo" }),
        makeSuggestion({ id: "s3", anchor: "mismo", context: "Contexto mismo" }),
        makeSuggestion({ id: "s4", anchor: "mismo", context: "Contexto mismo" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(1);
      expect(ctx.uniqueSuggestions?.[0].id).toBe("s1");
    });

    it("should still call next() even when many duplicates are removed", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", anchor: "dup", context: "Contexto dup" }),
        makeSuggestion({ id: "s2", anchor: "dup", context: "Contexto dup" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Single suggestion
  // -----------------------------------------------------------------------

  describe("single suggestion", () => {
    it("should pass through a single suggestion unchanged", async () => {
      const suggestion = makeSuggestion({ id: "only-one" });
      const ctx = makePipelineContext([suggestion]);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toEqual([suggestion]);
    });
  });

  // -----------------------------------------------------------------------
  // Empty input
  // -----------------------------------------------------------------------

  describe("empty input", () => {
    it("should set uniqueSuggestions to empty array when rawSuggestions is empty", async () => {
      const ctx = makePipelineContext([]);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toEqual([]);
      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Does NOT abort
  // -----------------------------------------------------------------------

  describe("abort behavior", () => {
    it("should never set ctx.aborted", async () => {
      const ctx = makePipelineContext([]);

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBeUndefined();
    });
  });

  // -----------------------------------------------------------------------
  // Deduplication only by context+anchor, not by other fields
  // -----------------------------------------------------------------------

  describe("deduplication key", () => {
    it("should not deduplicate suggestions with same suggestedText but different anchors", async () => {
      const suggestions = [
        makeSuggestion({
          id: "s1",
          anchor: "error uno",
          context: "Contexto error uno",
          suggestedText: "corrección",
        }),
        makeSuggestion({
          id: "s2",
          anchor: "error dos",
          context: "Contexto error dos",
          suggestedText: "corrección",
        }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(2);
    });

    it("should not deduplicate suggestions with same category but different anchors", async () => {
      const suggestions = [
        makeSuggestion({
          id: "s1",
          anchor: "alfa",
          context: "Contexto alfa",
          category: "Redundancia",
        }),
        makeSuggestion({
          id: "s2",
          anchor: "beta",
          context: "Contexto beta",
          category: "Redundancia",
        }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(2);
    });
  });

  // -----------------------------------------------------------------------
  // comment-only suggestions — deduplicated by exact semantic identity
  // -----------------------------------------------------------------------

  describe("comment-only deduplication behavior", () => {
    it("should keep all comment-only suggestions even when they share the same anchor", async () => {
      // Two comment-only suggestions targeting the same phrase are both valid:
      // they produce independent Word comments and never conflict with each other.
      const suggestions = [
        makeCommentOnlySuggestion({
          id: "c1",
          anchor: "misma frase",
          context: "Contexto misma frase",
        }),
        makeCommentOnlySuggestion({
          id: "c2",
          anchor: "misma frase",
          context: "Otro contexto misma frase",
        }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(2);
      expect(ctx.uniqueSuggestions?.map((s) => s.id)).toEqual(["c1", "c2"]);
    });

    it("should keep multiple comment-only suggestions from different runs that share anchor", async () => {
      const suggestions = [
        makeCommentOnlySuggestion({ id: "run1-c1", anchor: "texto", context: "Contexto run1" }),
        makeCommentOnlySuggestion({ id: "run2-c1", anchor: "texto", context: "Contexto run2 uno" }),
        makeCommentOnlySuggestion({ id: "run2-c2", anchor: "texto", context: "Contexto run2 dos" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(3);
    });

    it("should deduplicate track-change while keeping all comment-only when both share the same anchor", async () => {
      // track-change duplicates are still deduplicated; comment-only with different content are kept
      const suggestions = [
        makeSuggestion({ id: "tc1", anchor: "frase compartida", context: "Contexto compartido" }),
        makeSuggestion({ id: "tc2", anchor: "frase compartida", context: "Contexto compartido" }), // dup — removed
        makeCommentOnlySuggestion({
          id: "co1",
          anchor: "frase compartida",
          context: "Contexto comentario uno",
        }),
        makeCommentOnlySuggestion({
          id: "co2",
          anchor: "frase compartida",
          context: "Contexto comentario dos",
        }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      // tc1 kept, tc2 removed; co1 and co2 both kept
      expect(ctx.uniqueSuggestions).toHaveLength(3);
      expect(ctx.uniqueSuggestions?.map((s) => s.id)).toEqual(["tc1", "co1", "co2"]);
    });

    it("should deduplicate exact semantic comment-only duplicates even with different ids", async () => {
      const s1 = makeCommentOnlySuggestion({
        id: "chunk0-9",
        anchor: "shared",
        context: "Contexto shared",
        justification: "Misma observación",
        category: "estilo",
      });
      const s2 = makeCommentOnlySuggestion({
        id: "chunk1-9",
        anchor: "shared",
        context: "Contexto shared",
        justification: "Misma observación",
        category: "estilo",
      });
      const ctx = makePipelineContext([s1, s2]);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(1);
      expect(ctx.uniqueSuggestions?.[0].id).toBe("chunk0-9");
    });

    it("should keep comment-only suggestions on the same anchor when the comment differs", async () => {
      const suggestions = [
        makeCommentOnlySuggestion({
          id: "c1",
          anchor: "shared",
          context: "Contexto shared",
          justification: "Primera observación",
        }),
        makeCommentOnlySuggestion({
          id: "c2",
          anchor: "shared",
          context: "Contexto shared",
          justification: "Segunda observación",
        }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(2);
      expect(ctx.uniqueSuggestions?.map((s) => s.id)).toEqual(["c1", "c2"]);
    });
  });
});
