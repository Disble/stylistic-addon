import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { CommandResult, Suggestion } from "../../domain/types";
import {
  getCommandMocks,
  installWordWithContext,
  makeSuggestion,
} from "./WordAdapterTestHelper";
import { WordAdapter } from "./WordAdapter";

describe("WordAdapter.applySuggestions", () => {
  const commandMocks = getCommandMocks();
  let adapter: WordAdapter;

  beforeEach(() => {
    vi.clearAllMocks();
    adapter = new WordAdapter();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("returns the empty aggregate result when there are no suggestions", async () => {
    installWordWithContext({
      document: {
        contentControls: { load: vi.fn(), items: [] },
        load: vi.fn(),
        changeTrackingMode: "off",
      },
      sync: vi.fn().mockResolvedValue(undefined),
    });

    await expect(adapter.applySuggestions([])).resolves.toEqual({
      successCount: 0,
      failedSuggestions: [],
      pendingAfter: {
        pendingStylisticArtifacts: 0,
        hasPendingStylisticArtifacts: false,
        trackChangesActive: false,
      },
      documentState: "idle",
      trackChangesActivatedForBatch: false,
    });

    expect(commandMocks.constructor).not.toHaveBeenCalled();
  });

  it("does not enable Track Changes for comment-only batches", async () => {
    const context = {
      document: {
        contentControls: { load: vi.fn(), items: [{ tag: "stylistic:comment-only:s-1" }] },
        load: vi.fn(),
        changeTrackingMode: "off",
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    installWordWithContext(context);
    commandMocks.execute.mockResolvedValueOnce({ success: true, commandId: "s-1" });

    const result = await adapter.applySuggestions([
      makeSuggestion({ id: "s-1", type: "comment-only", suggestedText: undefined }),
    ]);

    expect(result.trackChangesActivatedForBatch).toBe(false);
    expect(context.document.changeTrackingMode).toBe("off");
  });

  it("enables Track Changes lazily for the first track-change batch", async () => {
    const context = {
      document: {
        contentControls: { load: vi.fn(), items: [{ tag: "stylistic:track-change:s-1" }] },
        load: vi.fn(),
        changeTrackingMode: "off",
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    installWordWithContext(context);
    commandMocks.execute.mockResolvedValueOnce({ success: true, commandId: "s-1" });

    const result = await adapter.applySuggestions([makeSuggestion({ id: "s-1" })]);

    expect(result.trackChangesActivatedForBatch).toBe(true);
    expect(result.pendingAfter.trackChangesActive).toBe(true);
    expect(context.document.changeTrackingMode).toBe("trackAll");
  });

  it("exposes the orchestrator aggregate result through the adapter facade", async () => {
    installWordWithContext({
      document: {
        contentControls: { load: vi.fn(), items: [{ tag: "stylistic:track-change:s-1" }] },
        load: vi.fn(),
        changeTrackingMode: "trackAll",
      },
      sync: vi.fn().mockResolvedValue(undefined),
    });
    const first = makeSuggestion({ id: "s-1", anchor: "uno", context: "Contexto uno" });
    const second = makeSuggestion({ id: "s-2", anchor: "dos", context: "Contexto dos" });

    commandMocks.execute
      .mockResolvedValueOnce({ success: true, commandId: "s-2" })
      .mockResolvedValueOnce({
        success: false,
        commandId: "s-1",
        error: "Texto original no encontrado",
      });

    await expect(adapter.applySuggestions([first, second])).resolves.toEqual({
      successCount: 1,
      failedSuggestions: [
        {
          suggestion: first,
          reason: "not-found",
          message: "Texto original no encontrado",
        },
      ],
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      trackChangesActivatedForBatch: false,
    });
  });

  it("rebuilds a localized reread hint from Word when the batch requests it", async () => {
    const paragraphRange = {
      text: "prefijo overlap-anchor sufijo",
      load: vi.fn(),
    };
    const localizedRange = {
      text: "prefijo overlap-anchor sufijo",
      load: vi.fn(),
      paragraphs: {
        getFirst: () => ({
          getRange: () => paragraphRange,
        }),
      },
      search: vi.fn().mockReturnValue({
        items: [{ text: "overlap-anchor", load: vi.fn() }],
        load: vi.fn(),
      }),
    };

    installWordWithContext({
      document: {
        contentControls: { load: vi.fn(), items: [] },
        load: vi.fn(),
        changeTrackingMode: "trackAll",
        body: {
          search: vi
            .fn()
            .mockReturnValueOnce({ items: [localizedRange], load: vi.fn() })
            .mockReturnValue({ items: [], load: vi.fn() }),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    });

    const rereadHint = await (
      adapter as unknown as {
        rereadSuggestionPositionHint: (
          suggestion: Suggestion,
          patch: NonNullable<CommandResult["mutationPatch"]>,
        ) => Promise<Suggestion["positionHint"] | undefined>;
      }
    ).rereadSuggestionPositionHint(
      makeSuggestion({ id: "s-overlap", anchor: "overlap-anchor" }),
      {
        suggestionId: "s-legacy",
        snapshotVersion: 1,
        originalText: "texto viejo",
        updatedText: "prefijo overlap-anchor sufijo",
        deltaLength: -7,
        affectedStart: 20,
        affectedEnd: 30,
      },
    );

    expect(rereadHint).toEqual({
      start: 8,
      end: 22,
      snapshotVersion: 1,
      source: "localized-reread",
    });
  });
});
