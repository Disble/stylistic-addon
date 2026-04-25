import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import {
  createRange,
  installWordContext,
  makeSuggestion,
} from "./ApplySuggestionCommandTestHelper";
import { WordTextLocatorAdapter } from "./WordTextLocatorAdapter";

const textLocator = new WordTextLocatorAdapter();

describe("ApplySuggestionCommand search behavior", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
    vi.spyOn(console, "warn").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("returns a distinct error when the surrounding context cannot be located", async () => {
    installWordContext({ contextSearchSequence: [[], [], []] });

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor no encontrado en el contexto",
    });
  });

  it("returns a distinct error when the anchor is missing inside the located context", async () => {
    installWordContext({
      contextSearchSequence: [[createRange({ text: "Contexto con texto original." })]],
      anchorSearchSequence: [[], [], []],
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor no encontrado en el contexto",
    });
  });

  it("uses body search for context and context search for anchor before replacing text", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

    expect(result).toMatchObject({ success: true, commandId: "s1" });
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

  it("aborts before mutation when the anchor is already covered by a content control", async () => {
    const env = installWordContext({
      anchorRangeParentCC: {
        tag: "stylistic:track-change:existing",
        isNullObject: false,
      },
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor cubierto por un Content Control existente",
    });
    expect(env.anchorRange.insertText).not.toHaveBeenCalled();
  });
});
