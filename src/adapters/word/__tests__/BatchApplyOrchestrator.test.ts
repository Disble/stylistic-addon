import { beforeEach, describe, expect, it, vi } from "vitest";
import type { CommandResult } from "../../../domain/DocumentApplication.types";
import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import { BatchApplyOrchestrator } from "../BatchApplyOrchestrator";

const hoistedCommandMocks = vi.hoisted(() => ({
  constructor: vi.fn<(suggestion: Suggestion) => void>(),
  execute: vi.fn<(suggestion: Suggestion) => Promise<CommandResult>>(),
}));

vi.mock("../ApplySuggestionCommand", () => ({
  ApplySuggestionCommand: class {
    private readonly suggestion: Suggestion;

    constructor(suggestion: Suggestion) {
      this.suggestion = suggestion;
      hoistedCommandMocks.constructor(suggestion);
    }

    execute() {
      return hoistedCommandMocks.execute(this.suggestion);
    }
  },
}));

describe("BatchApplyOrchestrator", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    hoistedCommandMocks.execute.mockResolvedValue({
      success: true,
      commandId: "ok",
    });
  });

  it("orders snapshot-comparable suggestions from later to earlier positions", async () => {
    const orchestrator = makeOrchestrator();
    const early = makeSuggestion("s-early", 20, 30);
    const middle = makeSuggestion("s-middle", 80, 90);
    const late = makeSuggestion("s-late", 120, 130);

    await orchestrator.run([middle, late, early]);

    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(1, late);
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(2, middle);
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(3, early);
  });

  it("marks overlapping comparable hints for local reread and rebases safe later hints", async () => {
    const orchestrator = makeOrchestrator();
    const applied = makeLegacySuggestion("s-applied");
    const overlapping = makeSuggestion("s-overlap", 45, 55);
    const safeLater = makeSuggestion("s-safe", 100, 110);

    hoistedCommandMocks.execute
      .mockResolvedValueOnce({
        success: true,
        commandId: "s-applied",
        mutationPatch: {
          suggestionId: "s-applied",
          snapshotVersion: 1,
          originalText: "abcdefghij",
          updatedText: "abc",
          deltaLength: -7,
          affectedStart: 40,
          affectedEnd: 50,
        },
      })
      .mockResolvedValueOnce({ success: true, commandId: "s-safe" })
      .mockResolvedValueOnce({ success: true, commandId: "s-overlap" });

    await orchestrator.run([overlapping, safeLater, applied]);

    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(1, applied);
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(2, {
      ...safeLater,
      positionHint: {
        start: 100,
        end: 110,
        snapshotVersion: 0,
        source: "snapshot",
      },
    });
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(3, {
      ...overlapping,
      positionHint: {
        start: 45,
        end: 55,
        snapshotVersion: 0,
        source: "snapshot",
      },
    });
  });

  it("uses localized reread when a reread-required hint cannot be reseeded from the latest patch", async () => {
    const rereadSuggestionPositionHint = vi.fn().mockResolvedValue({
      start: 210,
      end: 224,
      snapshotVersion: 1,
      source: "localized-reread",
    });
    const orchestrator = makeOrchestrator({ rereadSuggestionPositionHint });

    const overlapping: Suggestion = {
      ...makeSuggestion("s-overlap", 24, 49),
      anchor: "anchor-missing-from-patch",
      positionHint: {
        start: 24,
        end: 49,
        snapshotVersion: 0,
        source: "snapshot",
        requiresLocalReread: true,
      },
    };
    const safeLater = makeSuggestion("s-safe", 100, 110);
    const applied = makeLegacySuggestion("s-legacy");
    const latestPatch = {
      suggestionId: "s-legacy",
      snapshotVersion: 1,
      originalText: "012345678901234567890123unchanged tail",
      updatedText: "01234567890123456unchanged tail",
      deltaLength: -7,
      affectedStart: 20,
      affectedEnd: 30,
    };

    hoistedCommandMocks.execute
      .mockResolvedValueOnce({
        success: true,
        commandId: "s-legacy",
        mutationPatch: latestPatch,
      })
      .mockResolvedValueOnce({ success: true, commandId: "s-safe" })
      .mockResolvedValueOnce({ success: true, commandId: "s-overlap" });

    await orchestrator.run([overlapping, safeLater, applied]);

    expect(rereadSuggestionPositionHint).toHaveBeenCalledWith(
      expect.objectContaining({ id: "s-overlap" }),
      latestPatch
    );
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(3, {
      ...overlapping,
      positionHint: {
        start: 210,
        end: 224,
        snapshotVersion: 1,
        source: "localized-reread",
        requiresLocalReread: false,
      },
    });
  });

  it("activates Track Changes only once for the first track-change suggestion", async () => {
    const ensureTrackChangesActive = vi.fn().mockResolvedValue(true);
    const orchestrator = makeOrchestrator({ ensureTrackChangesActive });
    const commentOnly = makeSuggestion("s-comment", 10, 20, "comment-only");
    const trackA = makeSuggestion("s-track-a", 40, 50);
    const trackB = makeSuggestion("s-track-b", 80, 90);

    const result = await orchestrator.run([commentOnly, trackA, trackB]);

    expect(ensureTrackChangesActive).toHaveBeenCalledTimes(1);
    expect(result.trackChangesActivatedForBatch).toBe(true);
  });

  it("continues after failures, classifies reasons, and reports progress", async () => {
    const onProgress = vi.fn();
    const orchestrator = makeOrchestrator({
      getDocumentReviewState: vi
        .fn()
        .mockResolvedValueOnce({
          pendingStylisticArtifacts: 0,
          hasPendingStylisticArtifacts: false,
          trackChangesActive: true,
        })
        .mockResolvedValueOnce({
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        }),
      deriveDocumentState: vi.fn().mockReturnValue("pending-review"),
    });
    const notFound = makeLegacySuggestion("s-not-found");
    const covered = makeLegacySuggestion("s-covered");
    const success = makeLegacySuggestion("s-success");

    hoistedCommandMocks.execute
      .mockResolvedValueOnce({
        success: false,
        commandId: "s-success",
        error: "Texto original no encontrado",
      })
      .mockResolvedValueOnce({
        success: false,
        commandId: "s-covered",
        error: "Anchor dentro de content control existente",
      })
      .mockResolvedValueOnce({
        success: true,
        commandId: "s-not-found",
      });

    const result = await orchestrator.run([notFound, covered, success], onProgress);

    expect(result.successCount).toBe(1);
    expect(result.failedSuggestions).toEqual([
      {
        suggestion: success,
        reason: "not-found",
        message: "Texto original no encontrado",
      },
      {
        suggestion: covered,
        reason: "covered-by-existing-cc",
        message: "Anchor dentro de content control existente",
      },
    ]);
    expect(onProgress).toHaveBeenNthCalledWith(
      1,
      "applying",
      1,
      3,
      "Aplicando sugerencia 1 de 3..."
    );
    expect(onProgress).toHaveBeenNthCalledWith(
      3,
      "applying",
      3,
      3,
      "Aplicando sugerencia 3 de 3..."
    );
  });
});

function makeOrchestrator(
  overrides: Partial<{
    ensureTrackChangesActive: () => Promise<boolean>;
    getDocumentReviewState: () => Promise<{
      pendingStylisticArtifacts: number;
      hasPendingStylisticArtifacts: boolean;
      trackChangesActive: boolean;
    }>;
    deriveDocumentState: () => "idle" | "pending-review" | "ready-to-disable-track-changes";
    rereadSuggestionPositionHint: (
      suggestion: Suggestion,
      patch: NonNullable<CommandResult["mutationPatch"]>
    ) => Promise<Suggestion["positionHint"] | undefined>;
  }> = {}
): BatchApplyOrchestrator {
  return new BatchApplyOrchestrator({
    ensureTrackChangesActive: vi.fn().mockResolvedValue(false),
    getDocumentReviewState: vi.fn().mockResolvedValue({
      pendingStylisticArtifacts: 0,
      hasPendingStylisticArtifacts: false,
      trackChangesActive: true,
    }),
    deriveDocumentState: vi.fn().mockReturnValue("idle"),
    ...overrides,
  });
}

function makeSuggestion(
  id: string,
  start: number,
  end: number,
  type: Suggestion["type"] = "track-change"
): Suggestion {
  return {
    id,
    anchor: `anchor-${id}`,
    context: `context-${id}`,
    suggestedText: type === "track-change" ? `replacement-${id}` : undefined,
    justification: "justification",
    category: "category",
    severity: "medium",
    type,
    positionHint: {
      start,
      end,
      snapshotVersion: 0,
      source: "snapshot",
    },
  };
}

function makeLegacySuggestion(id: string): Suggestion {
  return {
    id,
    anchor: `anchor-${id}`,
    context: `context-${id}`,
    suggestedText: `replacement-${id}`,
    justification: "justification",
    category: "category",
    severity: "medium",
    type: "track-change",
  };
}
