import { DeduplicateHandler } from "./DeduplicateHandler";
import { PipelineContext } from "../PipelineContext";
import { PipelineEventEmitter } from "../PipelineEvents";
import type { IDocumentPort, IAnalysisPort } from "../../ports";
import type { Suggestion } from "../../types";

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

function makeMockPorts(): { documentPort: IDocumentPort; analysisPort: IAnalysisPort } {
  return {
    documentPort: {
      getTextToAnalyze: vi.fn(),
      getAppliedOriginalTexts: vi.fn(),
      applySuggestions: vi.fn(),
      cleanupResolvedComments: vi.fn(),
      acceptSuggestion: vi.fn(),
      rejectSuggestion: vi.fn(),
    },
    analysisPort: {
      checkConnection: vi.fn(),
      analyzeChunk: vi.fn(),
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
    profile: "general",
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
        makeSuggestion({ id: "s1", originalText: "alfa" }),
        makeSuggestion({ id: "s2", originalText: "beta" }),
        makeSuggestion({ id: "s3", originalText: "gamma" }),
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
    it("should remove exact duplicate originalText", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", originalText: "repetido" }),
        makeSuggestion({ id: "s2", originalText: "repetido" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(1);
      expect(ctx.uniqueSuggestions![0].id).toBe("s1");
    });

    it("should keep the first occurrence and discard subsequent duplicates", async () => {
      const suggestions = [
        makeSuggestion({ id: "first", originalText: "duplicado", suggestedText: "versión 1" }),
        makeSuggestion({ id: "second", originalText: "duplicado", suggestedText: "versión 2" }),
        makeSuggestion({ id: "third", originalText: "duplicado", suggestedText: "versión 3" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(1);
      expect(ctx.uniqueSuggestions![0].id).toBe("first");
      expect(ctx.uniqueSuggestions![0].suggestedText).toBe("versión 1");
    });

    it("should handle a mix of unique and duplicate suggestions", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", originalText: "alfa" }),
        makeSuggestion({ id: "s2", originalText: "beta" }),
        makeSuggestion({ id: "s3", originalText: "alfa" }), // dup of s1
        makeSuggestion({ id: "s4", originalText: "gamma" }),
        makeSuggestion({ id: "s5", originalText: "beta" }), // dup of s2
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(3);
      expect(ctx.uniqueSuggestions!.map((s) => s.id)).toEqual(["s1", "s2", "s4"]);
    });
  });

  // -----------------------------------------------------------------------
  // Case-insensitive comparison
  // -----------------------------------------------------------------------

  describe("case-insensitive deduplication", () => {
    it("should treat different cases as duplicates", async () => {
      const suggestions = [
        makeSuggestion({ id: "lower", originalText: "texto" }),
        makeSuggestion({ id: "upper", originalText: "TEXTO" }),
        makeSuggestion({ id: "mixed", originalText: "Texto" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(1);
      expect(ctx.uniqueSuggestions![0].id).toBe("lower");
    });

    it("should preserve the original casing of the kept suggestion", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", originalText: "MiTexto" }),
        makeSuggestion({ id: "s2", originalText: "mitexto" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions![0].originalText).toBe("MiTexto");
    });
  });

  // -----------------------------------------------------------------------
  // All duplicates
  // -----------------------------------------------------------------------

  describe("all duplicates", () => {
    it("should keep exactly one when all suggestions are duplicates", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", originalText: "mismo" }),
        makeSuggestion({ id: "s2", originalText: "mismo" }),
        makeSuggestion({ id: "s3", originalText: "mismo" }),
        makeSuggestion({ id: "s4", originalText: "mismo" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(1);
      expect(ctx.uniqueSuggestions![0].id).toBe("s1");
    });

    it("should still call next() even when many duplicates are removed", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", originalText: "dup" }),
        makeSuggestion({ id: "s2", originalText: "dup" }),
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
  // Deduplication only by originalText, not by other fields
  // -----------------------------------------------------------------------

  describe("deduplication key", () => {
    it("should not deduplicate suggestions with same suggestedText but different originalText", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", originalText: "error uno", suggestedText: "corrección" }),
        makeSuggestion({ id: "s2", originalText: "error dos", suggestedText: "corrección" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(2);
    });

    it("should not deduplicate suggestions with same category but different originalText", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", originalText: "alfa", category: "Redundancia" }),
        makeSuggestion({ id: "s2", originalText: "beta", category: "Redundancia" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.uniqueSuggestions).toHaveLength(2);
    });
  });
});
