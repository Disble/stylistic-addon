import { beforeEach, describe, expect, it, vi } from "vitest";
import { BatchApplyOrchestrator } from "./BatchApplyOrchestrator";
import type { CommandResult, Suggestion } from "../../domain/types";

const hoistedCommandMocks = vi.hoisted(() => ({
  constructor: vi.fn<(suggestion: Suggestion) => void>(),
  execute: vi.fn<(suggestion: Suggestion) => Promise<CommandResult>>(),
}));

vi.mock("./ApplySuggestionCommand", () => ({
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

  it("prefers snapshot position hints over raw backend array order", async () => {
    const orchestrator = new BatchApplyOrchestrator({
      ensureTrackChangesActive: vi.fn().mockResolvedValue(false),
      getDocumentReviewState: vi.fn().mockResolvedValue({
        pendingStylisticArtifacts: 0,
        hasPendingStylisticArtifacts: false,
        trackChangesActive: true,
      }),
      deriveDocumentState: vi.fn().mockReturnValue("idle"),
    });

    const backendFirst = makeSuggestion("s-a", 20, 30);
    const backendSecond = makeSuggestion("s-b", 80, 90);
    const backendThird = makeSuggestion("s-c", 120, 130);

    await orchestrator.run([
      backendSecond,
      backendThird,
      backendFirst,
    ]);

    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(1, backendThird);
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(2, backendSecond);
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(3, backendFirst);
  });

  it("rebases pending snapshot position hints after a successful mutation patch in legacy-order batches", async () => {
    const orchestrator = new BatchApplyOrchestrator({
      ensureTrackChangesActive: vi.fn().mockResolvedValue(false),
      getDocumentReviewState: vi.fn().mockResolvedValue({
        pendingStylisticArtifacts: 0,
        hasPendingStylisticArtifacts: false,
        trackChangesActive: true,
      }),
      deriveDocumentState: vi.fn().mockReturnValue("idle"),
    });

    const hintedLater = makeSuggestion("s-later", 100, 110);
    const hintedEvenLater = makeSuggestion("s-later-2", 130, 140);
    const legacyFirst = makeLegacySuggestion("s-legacy-first");

    hoistedCommandMocks.execute
      .mockResolvedValueOnce({
        success: true,
        commandId: "s-legacy-first",
        mutationPatch: {
          suggestionId: "s-legacy-first",
          originalText: "abcdefghij",
          updatedText: "abc",
          deltaLength: -7,
          affectedStart: 20,
          affectedEnd: 30,
        },
      })
      .mockResolvedValueOnce({
        success: true,
        commandId: "s-later-2",
      })
      .mockResolvedValueOnce({
        success: true,
        commandId: "s-later",
      });

    await orchestrator.run([hintedLater, hintedEvenLater, legacyFirst]);

    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(1, legacyFirst);
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(2, {
      ...hintedEvenLater,
      positionHint: {
        start: 123,
        end: 133,
        source: "snapshot",
      },
    });
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(3, {
      ...hintedLater,
      positionHint: {
        start: 93,
        end: 103,
        source: "snapshot",
      },
    });
  });

  it("marks overlapping snapshot hints for local reread and lets safe hinted suggestions run first", async () => {
    const orchestrator = new BatchApplyOrchestrator({
      ensureTrackChangesActive: vi.fn().mockResolvedValue(false),
      getDocumentReviewState: vi.fn().mockResolvedValue({
        pendingStylisticArtifacts: 0,
        hasPendingStylisticArtifacts: false,
        trackChangesActive: true,
      }),
      deriveDocumentState: vi.fn().mockReturnValue("idle"),
    });

    const overlappingHint = makeSuggestion("s-overlap", 24, 34);
    const safeLaterHint = makeSuggestion("s-safe", 100, 110);
    const legacyFirst = makeLegacySuggestion("s-legacy-first");

    hoistedCommandMocks.execute
      .mockResolvedValueOnce({
        success: true,
        commandId: "s-legacy-first",
        mutationPatch: {
          suggestionId: "s-legacy-first",
          originalText: "abcdefghij",
          updatedText: "abc",
          deltaLength: -7,
          affectedStart: 20,
          affectedEnd: 30,
        },
      })
      .mockResolvedValueOnce({
        success: true,
        commandId: "s-safe",
      })
      .mockResolvedValueOnce({
        success: true,
        commandId: "s-overlap",
      });

    await orchestrator.run([overlappingHint, safeLaterHint, legacyFirst]);

    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(1, legacyFirst);
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(2, {
      ...safeLaterHint,
      positionHint: {
        start: 93,
        end: 103,
        source: "snapshot",
      },
    });
    expect(hoistedCommandMocks.constructor).toHaveBeenNthCalledWith(3, {
      ...overlappingHint,
      positionHint: {
        start: 24,
        end: 34,
        source: "snapshot",
        requiresLocalReread: true,
      },
    });
  });
});

/** Builds a suggestion fixture with explicit snapshot-position hints. */
function makeSuggestion(id: string, start: number, end: number): Suggestion {
  return {
    id,
    anchor: `anchor-${id}`,
    context: `context-${id}`,
    suggestedText: `replacement-${id}`,
    justification: "justification",
    category: "category",
    severity: "medium",
    type: "track-change",
    positionHint: {
      start,
      end,
      source: "snapshot",
    },
  };
}

/** Builds a legacy suggestion fixture without snapshot-position hints. */
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
