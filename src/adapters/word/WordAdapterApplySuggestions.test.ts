import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  getCommandMocks,
  installWordWithContext,
  makeSuggestion,
} from "./WordAdapterTestHelper";
import { WordAdapter } from "./WordAdapter";
import type { CommandResult, Suggestion } from "../../domain/types";

describe("WordAdapter.applySuggestions", () => {
  const commandMocks = getCommandMocks();
  let adapter: WordAdapter;
  let logSpy: ReturnType<typeof vi.spyOn>;
  let warnSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.clearAllMocks();
    adapter = new WordAdapter();
    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
    warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
  });

  afterEach(() => {
    logSpy.mockRestore();
    warnSpy.mockRestore();
    vi.unstubAllGlobals();
  });

  it("returns an empty insertion result when there are no suggestions", async () => {
    installWordWithContext({
      document: {
        contentControls: { load: vi.fn(), items: [] },
        load: vi.fn(),
        changeTrackingMode: "off",
      },
      sync: vi.fn().mockResolvedValue(undefined),
    });
    const onProgress = vi.fn();

    await expect(adapter.applySuggestions([], onProgress)).resolves.toEqual({
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
    expect(commandMocks.execute).not.toHaveBeenCalled();
    expect(onProgress).not.toHaveBeenCalled();
  });

  it("maps command results into successCount, failedSuggestions, and progress events", async () => {
    installWordWithContext({
      document: {
        contentControls: { load: vi.fn(), items: [{ tag: "stylistic:track-change:s-2" }] },
        load: vi.fn(),
        changeTrackingMode: "trackAll",
      },
      sync: vi.fn().mockResolvedValue(undefined),
    });
    const first = makeSuggestion({ id: "s-1", anchor: "uno", context: "Contexto uno" });
    const second = makeSuggestion({ id: "s-2", anchor: "dos", context: "Contexto dos" });
    const onProgress = vi.fn();

    commandMocks.execute
      .mockResolvedValueOnce({ success: true, commandId: "s-2" })
      .mockResolvedValueOnce({
        success: false,
        commandId: "s-1",
        error: "Texto original no encontrado",
      });

    await expect(
      adapter.applySuggestions([first, second], onProgress),
    ).resolves.toEqual({
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

    expect(commandMocks.constructor).toHaveBeenNthCalledWith(1, second);
    expect(commandMocks.constructor).toHaveBeenNthCalledWith(2, first);
    expect(commandMocks.execute).toHaveBeenNthCalledWith(1, second);
    expect(commandMocks.execute).toHaveBeenNthCalledWith(2, first);
    expect(onProgress).toHaveBeenNthCalledWith(
      1,
      "applying",
      1,
      2,
      "Aplicando sugerencia 1 de 2...",
    );
    expect(onProgress).toHaveBeenNthCalledWith(
      2,
      "applying",
      2,
      2,
      "Aplicando sugerencia 2 de 2...",
    );
    expect(warnSpy).toHaveBeenCalledOnce();
  });

  it("converts command rejections into failed suggestions and continues processing", async () => {
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
    const third = makeSuggestion({ id: "s-3", anchor: "tres", context: "Contexto tres" });
    const onProgress = vi.fn();

    commandMocks.execute
      .mockResolvedValueOnce({ success: true, commandId: "s-3" })
      .mockRejectedValueOnce(new Error("insert failed"))
      .mockResolvedValueOnce({ success: true, commandId: "s-1" });

    await expect(
      adapter.applySuggestions([first, second, third], onProgress),
    ).resolves.toEqual({
      successCount: 2,
      failedSuggestions: [
        {
          suggestion: second,
          reason: "command-error",
          message: "insert failed",
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

    expect(commandMocks.constructor).toHaveBeenNthCalledWith(1, third);
    expect(commandMocks.constructor).toHaveBeenNthCalledWith(2, second);
    expect(commandMocks.constructor).toHaveBeenNthCalledWith(3, first);
    expect(commandMocks.execute).toHaveBeenNthCalledWith(1, third);
    expect(commandMocks.execute).toHaveBeenNthCalledWith(2, second);
    expect(commandMocks.execute).toHaveBeenNthCalledWith(3, first);
    expect(warnSpy).toHaveBeenCalledOnce();
  });

  it("applies suggestions in reverse array order to avoid content-control interference", async () => {
    installWordWithContext({
      document: {
        contentControls: { load: vi.fn(), items: [{ tag: "stylistic:track-change:s-a" }] },
        load: vi.fn(),
        changeTrackingMode: "trackAll",
      },
      sync: vi.fn().mockResolvedValue(undefined),
    });
    const suggA = makeSuggestion({
      id: "s-a",
      anchor: "texto al inicio del doc",
      context: "Contexto texto al inicio del doc",
    });
    const suggB = makeSuggestion({
      id: "s-b",
      anchor: "texto al final del doc",
      context: "Contexto texto al final del doc",
    });
    const suggC = makeSuggestion({
      id: "s-c",
      anchor: "texto en el medio del doc",
      context: "Contexto texto en el medio del doc",
    });

    commandMocks.execute
      .mockResolvedValueOnce({ success: true, commandId: "s-c" })
      .mockResolvedValueOnce({ success: true, commandId: "s-b" })
      .mockResolvedValueOnce({ success: true, commandId: "s-a" });

    await adapter.applySuggestions([suggA, suggB, suggC]);

    expect(commandMocks.constructor).toHaveBeenNthCalledWith(1, suggC);
    expect(commandMocks.constructor).toHaveBeenNthCalledWith(2, suggB);
    expect(commandMocks.constructor).toHaveBeenNthCalledWith(3, suggA);
  });

  it("keeps same-paragraph suggestions independent by applying the later one first", async () => {
    installWordWithContext({
      document: {
        contentControls: { load: vi.fn(), items: [{ tag: "stylistic:track-change:s-late" }] },
        load: vi.fn(),
        changeTrackingMode: "trackAll",
      },
      sync: vi.fn().mockResolvedValue(undefined),
    });
    const earlyInParagraph = makeSuggestion({
      id: "s-early",
      anchor: "apenas eran un par de pasos alrededor de ella",
      context: "Contexto apenas eran un par de pasos alrededor de ella",
    });
    const lateInParagraph = makeSuggestion({
      id: "s-late",
      anchor: "asumió que eso sucedía porque perdía el control",
      context: "Contexto asumió que eso sucedía porque perdía el control",
    });

    commandMocks.execute
      .mockResolvedValueOnce({ success: true, commandId: "s-late" })
      .mockResolvedValueOnce({ success: true, commandId: "s-early" });

    const result = await adapter.applySuggestions([
      earlyInParagraph,
      lateInParagraph,
    ]);

    expect(result).toEqual({
      successCount: 2,
      failedSuggestions: [],
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      trackChangesActivatedForBatch: false,
    });
    expect(commandMocks.constructor).toHaveBeenNthCalledWith(1, lateInParagraph);
    expect(commandMocks.constructor).toHaveBeenNthCalledWith(2, earlyInParagraph);
  });

  it("does not enable Track Changes for comment-only-only batches", async () => {
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

  it("enables Track Changes lazily once for the first track-change suggestion in an off document", async () => {
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
    expect(context.document.changeTrackingMode).toBe("trackAll");
    expect(result.pendingAfter.trackChangesActive).toBe(true);
  });

  it("wires a localized reread strategy into the batch orchestrator", async () => {
    installWordWithContext({
      document: {
        contentControls: { load: vi.fn(), items: [] },
        load: vi.fn(),
        changeTrackingMode: "trackAll",
        body: {
          search: vi.fn().mockReturnValue({ items: [], load: vi.fn() }),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    });

    const adapter = new WordAdapter();
    commandMocks.execute.mockResolvedValueOnce({ success: true, commandId: "s-1" });

    const result = await adapter.applySuggestions([makeSuggestion({ id: "s-1" })]);

    expect(result).toEqual({
      successCount: 1,
      failedSuggestions: [],
      pendingAfter: {
        pendingStylisticArtifacts: 0,
        hasPendingStylisticArtifacts: false,
        trackChangesActive: true,
      },
      documentState: "ready-to-disable-track-changes",
      trackChangesActivatedForBatch: false,
    });
    expect(commandMocks.constructor).toHaveBeenCalledOnce();
    expect(commandMocks.execute).toHaveBeenCalledOnce();
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
    const bodySearch = vi
      .fn()
      .mockReturnValueOnce({
        items: [localizedRange],
        load: vi.fn(),
      })
      .mockReturnValue({ items: [], load: vi.fn() });

    const context = {
      document: {
        contentControls: { load: vi.fn(), items: [] },
        load: vi.fn(),
        changeTrackingMode: "trackAll",
        body: { search: bodySearch },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    installWordWithContext(context);

    const adapter = new WordAdapter();
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
