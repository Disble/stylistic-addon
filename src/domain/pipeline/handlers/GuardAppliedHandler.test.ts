import type { IAnalysisPort, IDocumentPort } from "../../ports";
import type { Suggestion } from "../../types";
import type { PipelineContext } from "../PipelineContext";
import { PipelineEventEmitter } from "../PipelineEvents";
import { GuardAppliedHandler } from "./GuardAppliedHandler";

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

function makeCommentOnlySuggestion(
  overrides: Partial<Suggestion> = {},
): Suggestion {
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

function makeMockDocumentPort(
  appliedTexts: Set<string> = new Set(),
): IDocumentPort {
  return {
    getTextToAnalyze: vi.fn(),
    getAppliedOriginalTexts: vi.fn().mockResolvedValue(appliedTexts),
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

function makeMockAnalysisPort(): IAnalysisPort {
  return {
    checkConnection: vi.fn(),
    submitChunkAnalysis: vi.fn(),
    pollChunkAnalysis: vi.fn(),
  };
}

function makePipelineContext(
  uniqueSuggestions: Suggestion[],
  appliedTexts: Set<string> = new Set(),
  overrides: Partial<PipelineContext> = {},
): PipelineContext {
  return {
    documentPort: makeMockDocumentPort(appliedTexts),
    analysisPort: makeMockAnalysisPort(),
    emitter: new PipelineEventEmitter(),
    genero: "general",
    maxChunkSize: 100_000,
    uniqueSuggestions,
    ...overrides,
  } as PipelineContext;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("GuardAppliedHandler", () => {
  let handler: GuardAppliedHandler;
  let next: () => Promise<void>;

  beforeEach(() => {
    handler = new GuardAppliedHandler();
    next = vi.fn<() => Promise<void>>();
  });

  // -----------------------------------------------------------------------
  // Happy path — no previously applied suggestions
  // -----------------------------------------------------------------------

  describe("no previously applied suggestions", () => {
    it("should pass all suggestions through as pendingSuggestions", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", anchor: "alfa", context: "Contexto alfa" }),
        makeSuggestion({ id: "s2", anchor: "beta", context: "Contexto beta" }),
      ];
      const ctx = makePipelineContext(suggestions);

      await handler.handle(ctx, next);

      expect(ctx.pendingSuggestions).toEqual(suggestions);
    });

    it("should call next() when there are pending suggestions", async () => {
      const ctx = makePipelineContext([makeSuggestion()]);

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });

    it("should call documentPort.getAppliedOriginalTexts()", async () => {
      const ctx = makePipelineContext([makeSuggestion()]);

      await handler.handle(ctx, next);

      expect(ctx.documentPort.getAppliedOriginalTexts).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Guard blocks — some suggestions already applied
  // -----------------------------------------------------------------------

  describe("some suggestions already applied", () => {
    it("should filter out suggestions whose anchor is in the applied set", async () => {
      const suggestions = [
        makeSuggestion({ id: "s1", anchor: "already applied", context: "Contexto already applied" }),
        makeSuggestion({ id: "s2", anchor: "new suggestion", context: "Contexto new suggestion" }),
        makeSuggestion({ id: "s3", anchor: "also applied", context: "Contexto also applied" }),
      ];
      const appliedTexts = new Set(["already applied", "also applied"]);
      const ctx = makePipelineContext(suggestions, appliedTexts);

      await handler.handle(ctx, next);

      expect(ctx.pendingSuggestions).toHaveLength(1);
      expect(ctx.pendingSuggestions?.[0].id).toBe("s2");
    });

    it("should call next() when at least one suggestion remains", async () => {
        const suggestions = [
        makeSuggestion({ id: "s1", anchor: "applied", context: "Contexto applied" }),
        makeSuggestion({ id: "s2", anchor: "pending", context: "Contexto pending" }),
      ];
      const ctx = makePipelineContext(suggestions, new Set(["applied"]));

      await handler.handle(ctx, next);

      expect(next).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Guard blocks — ALL suggestions already applied
  // -----------------------------------------------------------------------

  describe("all suggestions already applied", () => {
    it("should abort with 'already applied' message when all are filtered", async () => {
        const suggestions = [
        makeSuggestion({ id: "s1", anchor: "applied one", context: "Contexto applied one" }),
        makeSuggestion({ id: "s2", anchor: "applied two", context: "Contexto applied two" }),
      ];
      const appliedTexts = new Set(["applied one", "applied two"]);
      const ctx = makePipelineContext(suggestions, appliedTexts);

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBe(true);
      expect(ctx.abortReason).toBe(
        "Todas las sugerencias ya están aplicadas en el documento.",
      );
    });

    it("should emit abort event when all suggestions are already applied", async () => {
      const emitter = new PipelineEventEmitter();
      const emitAbort = vi.spyOn(emitter, "emitAbort");

      const suggestions = [makeSuggestion({ anchor: "applied", context: "Contexto applied" })];
      const ctx = makePipelineContext(suggestions, new Set(["applied"]), {
        emitter,
      });

      await handler.handle(ctx, next);

      expect(emitAbort).toHaveBeenCalledWith(
        "Todas las sugerencias ya están aplicadas en el documento.",
      );
    });

    it("should NOT call next() when aborting", async () => {
      const suggestions = [makeSuggestion({ anchor: "applied", context: "Contexto applied" })];
      const ctx = makePipelineContext(suggestions, new Set(["applied"]));

      await handler.handle(ctx, next);

      expect(next).not.toHaveBeenCalled();
    });
  });

  // -----------------------------------------------------------------------
  // Empty uniqueSuggestions — no skipped, but still zero pending
  // -----------------------------------------------------------------------

  describe("empty uniqueSuggestions input", () => {
    it("should abort with 'no suggestions' message when input is empty and nothing was skipped", async () => {
      const ctx = makePipelineContext([]);

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBe(true);
      expect(ctx.abortReason).toBe(
        "No se encontraron sugerencias editoriales.",
      );
    });

    it("should NOT call next() when input is empty", async () => {
      const ctx = makePipelineContext([]);

      await handler.handle(ctx, next);

      expect(next).not.toHaveBeenCalled();
    });
  });

  // -----------------------------------------------------------------------
  // Case-sensitive matching (guard uses exact anchor match via Set.has)
  // -----------------------------------------------------------------------

  describe("case-sensitive guard matching", () => {
    it("should NOT filter a suggestion when applied text differs in case", async () => {
      const suggestions = [makeSuggestion({ anchor: "Texto Original", context: "Contexto Texto Original" })];
      const appliedTexts = new Set(["texto original"]); // different case
      const ctx = makePipelineContext(suggestions, appliedTexts);

      await handler.handle(ctx, next);

      // The guard uses exact Set.has — case matters
      expect(ctx.pendingSuggestions).toHaveLength(1);
      expect(next).toHaveBeenCalledOnce();
    });

    it("should filter when casing matches exactly", async () => {
      const suggestions = [makeSuggestion({ anchor: "Texto Original", context: "Contexto Texto Original" })];
      const appliedTexts = new Set(["Texto Original"]); // exact match
      const ctx = makePipelineContext(suggestions, appliedTexts);

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBe(true);
    });
  });

  // -----------------------------------------------------------------------
  // Abort reason differentiates: skipped > 0 vs skipped === 0
  // -----------------------------------------------------------------------

  describe("abort reason differentiation", () => {
    it("should say 'already applied' when skipped > 0 and pending === 0", async () => {
      const suggestions = [makeSuggestion({ anchor: "ya aplicado", context: "Contexto ya aplicado" })];
      const ctx = makePipelineContext(suggestions, new Set(["ya aplicado"]));

      await handler.handle(ctx, next);

      expect(ctx.abortReason).toBe(
        "Todas las sugerencias ya están aplicadas en el documento.",
      );
    });

    it("should say 'no suggestions' when skipped === 0 and pending === 0", async () => {
      const ctx = makePipelineContext([]);

      await handler.handle(ctx, next);

      expect(ctx.abortReason).toBe(
        "No se encontraron sugerencias editoriales.",
      );
    });
  });

  // -----------------------------------------------------------------------
  // comment-only guard behavior
  // Phase 4 extends getAppliedOriginalTexts() to return comment-only CCs'
  // anchor texts, so the guard works for both types via the same Set.has()
  // check — no type-specific branching is needed here.
  // -----------------------------------------------------------------------

  describe("comment-only guard behavior", () => {
    it("should pass through a comment-only suggestion when its anchor is not in the applied set", async () => {
      const suggestions = [
        makeCommentOnlySuggestion({ id: "c1", anchor: "nueva frase", context: "Contexto nueva frase" }),
      ];
      const ctx = makePipelineContext(suggestions, new Set());

      await handler.handle(ctx, next);

      expect(ctx.pendingSuggestions).toHaveLength(1);
      expect(ctx.pendingSuggestions?.[0].id).toBe("c1");
      expect(next).toHaveBeenCalledOnce();
    });

    it("should filter out a comment-only suggestion whose anchor is already in the applied set", async () => {
      // getAppliedOriginalTexts() (extended in Phase 4) includes comment-only CC texts
      const suggestions = [
        makeCommentOnlySuggestion({
          id: "c1",
          anchor: "frase ya comentada",
          context: "Contexto frase ya comentada",
        }),
      ];
      const appliedTexts = new Set(["frase ya comentada"]);
      const ctx = makePipelineContext(suggestions, appliedTexts);

      await handler.handle(ctx, next);

      expect(ctx.aborted).toBe(true);
      expect(ctx.abortReason).toBe(
        "Todas las sugerencias ya están aplicadas en el documento.",
      );
    });

    it("should handle a mix of track-change and comment-only, filtering already-applied ones of both types", async () => {
      const suggestions = [
        makeSuggestion({ id: "tc1", anchor: "tc applied", context: "Contexto tc applied" }), // filtered
        makeSuggestion({ id: "tc2", anchor: "tc pending", context: "Contexto tc pending" }), // kept
        makeCommentOnlySuggestion({ id: "co1", anchor: "co applied", context: "Contexto co applied" }), // filtered
        makeCommentOnlySuggestion({ id: "co2", anchor: "co pending", context: "Contexto co pending" }), // kept
      ];
      const appliedTexts = new Set(["tc applied", "co applied"]);
      const ctx = makePipelineContext(suggestions, appliedTexts);

      await handler.handle(ctx, next);

      expect(ctx.pendingSuggestions).toHaveLength(2);
      expect(ctx.pendingSuggestions?.map((s) => s.id)).toEqual(["tc2", "co2"]);
      expect(next).toHaveBeenCalledOnce();
    });
  });
});
