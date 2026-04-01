import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  getCommandMocks,
  makeSuggestion,
} from "./WordAdapterTestHelper";
import { WordAdapter } from "./WordAdapter";

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
  });

  it("returns an empty insertion result when there are no suggestions", async () => {
    const onProgress = vi.fn();

    await expect(adapter.applySuggestions([], onProgress)).resolves.toEqual({
      successCount: 0,
      failedSuggestions: [],
    });

    expect(commandMocks.constructor).not.toHaveBeenCalled();
    expect(commandMocks.execute).not.toHaveBeenCalled();
    expect(onProgress).not.toHaveBeenCalled();
  });

  it("maps command results into successCount, failedSuggestions, and progress events", async () => {
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
      failedSuggestions: [first],
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
      failedSuggestions: [second],
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

    expect(result).toEqual({ successCount: 2, failedSuggestions: [] });
    expect(commandMocks.constructor).toHaveBeenNthCalledWith(1, lateInParagraph);
    expect(commandMocks.constructor).toHaveBeenNthCalledWith(2, earlyInParagraph);
  });
});
