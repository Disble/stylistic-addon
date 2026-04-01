import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import {
  createRange,
  installWordContext,
  makeSuggestion,
} from "./ApplySuggestionCommandTestHelper";

describe("ApplySuggestionCommand search behavior", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
    vi.spyOn(console, "warn").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("returns a distinct error when context is not found", async () => {
    installWordContext({ contextSearchSequence: [[], [], []] });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor no encontrado en el contexto",
    });
  });

  it("returns a distinct error when anchor is not found inside the located context", async () => {
    installWordContext({
      contextSearchSequence: [[createRange({ text: "Contexto con texto original." })]],
      anchorSearchSequence: [[], [], []],
    });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor no encontrado en el contexto",
    });
  });

  it("uses two-step search: body finds context, then context range finds anchor", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.context.document.body.search).toHaveBeenCalledWith(
      "Contexto con texto original.",
      { matchCase: true, matchWholeWord: false },
    );
    expect(env.bodyRange.search).toHaveBeenCalledWith("texto original", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "texto sugerido",
      "Replace",
    );
  });

  it("tries ignorePunct+ignoreSpace when exact anchor search fails", async () => {
    const env = installWordContext();
    env.bodyRange.search = vi.fn()
      .mockReturnValueOnce({ items: [], load: vi.fn() })
      .mockReturnValueOnce({ items: [env.anchorRange], load: vi.fn() });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.bodyRange.search).toHaveBeenNthCalledWith(1, "texto original", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(env.bodyRange.search).toHaveBeenNthCalledWith(2, "texto original", {
      matchCase: true,
      matchWholeWord: false,
      ignorePunct: true,
      ignoreSpace: true,
    });
  });

  it("skips the exact search when context text exceeds Word's 256-char search limit", async () => {
    const longContext = `Prefijo ${"x".repeat(270)}`;
    const env = installWordContext({
      documentText: longContext,
      contextText: longContext,
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({ context: longContext }),
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.context.document.body.search).toHaveBeenCalledTimes(1);
    expect(env.context.document.body.search).toHaveBeenCalledWith(longContext, {
      matchCase: true,
      matchWholeWord: false,
      ignorePunct: true,
      ignoreSpace: true,
    });
  });

  it("falls back to a whitespace-insensitive slice when backend and document spacing differ", async () => {
    const fallbackAnchor = createRange({ text: "texto\n\noriginal" });
    const env = installWordContext({
      contextText: "Contexto con texto\n\noriginal.",
      anchorSearchSequence: [[], [], [fallbackAnchor]],
    });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.bodyRange.load).toHaveBeenCalledWith("text");
    expect(env.bodyRange.search).toHaveBeenNthCalledWith(3, "texto\n\noriginal", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(fallbackAnchor.insertText).toHaveBeenCalledWith(
      "texto sugerido",
      "Replace",
    );
  });
});
