import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import { WordTextLocatorAdapter } from "./WordTextLocatorAdapter";
import {
  createRange,
  installWordContext,
  makeSuggestion,
  type ParentCC,
} from "./ApplySuggestionCommandTestHelper";

const IDENTITY_TITLE_PREFIX = "stylistic-meta-v2:";
const textLocator = new WordTextLocatorAdapter();

describe("ApplySuggestionCommand content-control recovery", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
    vi.spyOn(console, "warn").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("supports comment-only suggestions through the shared anchor resolver", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion({ type: "comment-only", suggestedText: undefined }),
      textLocator,
    ).execute();

    expect(result).toMatchObject({ success: true, commandId: "s1" });
    expect(env.anchorRange.insertComment).toHaveBeenCalledWith(
      "[Estilo]\nMejora la claridad",
    );
    expect(env.anchorRange.insertContentControl).toHaveBeenCalledOnce();
  });

  it("re-resolves the anchor after removing an existing stylistic content control", async () => {
    const coveredParentCC: ParentCC = {
      tag: "stylistic:track-change:s1",
      isNullObject: false,
      load: vi.fn(),
      delete: vi.fn(),
    };
    const coveredAnchor = createRange({
      text: "texto original",
      parentCC: coveredParentCC,
    });
    const coveredContext = createRange({
      text: "Contexto con texto original.",
      searchSequence: [[coveredAnchor]],
    });
    const freshAnchor = createRange({ text: "texto original" });
    const freshContext = createRange({
      text: "Contexto con texto original.",
      searchSequence: [[freshAnchor]],
    });

    const env = installWordContext({
      contextSearchSequence: [[coveredContext], [freshContext]],
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

    expect(result).toMatchObject({ success: true, commandId: "s1" });
    expect(coveredParentCC.delete).toHaveBeenCalledWith(true);
    expect(env.context.document.body.search).toHaveBeenCalledTimes(2);
    expect(freshAnchor.insertText).toHaveBeenCalled();
  });

  it("re-resolves the anchor after removing a chunk wrapper that still covers the text", async () => {
    const coveredParentCC: ParentCC = {
      tag: "chunk0-0",
      isNullObject: false,
      load: vi.fn(),
      delete: vi.fn(),
    };
    const coveredAnchor = createRange({
      text: "texto original",
      parentCC: coveredParentCC,
    });
    const coveredContext = createRange({
      text: "Contexto con texto original.",
      searchSequence: [[coveredAnchor]],
    });
    const freshAnchor = createRange({ text: "texto original" });
    const freshContext = createRange({
      text: "Contexto con texto original.",
      searchSequence: [[freshAnchor]],
    });

    installWordContext({
      contextSearchSequence: [[coveredContext], [freshContext]],
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion(),
      textLocator,
    ).execute();

    expect(result).toMatchObject({ success: true, commandId: "s1" });
    expect(coveredParentCC.delete).toHaveBeenCalledWith(true);
    expect(freshAnchor.insertText).toHaveBeenCalledOnce();
  });

  it("persists compound v2 replace metadata for replace suggestions", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "replace-1",
        anchor: "texto original",
        suggestedText: "texto sugerido",
        context: "Contexto con texto original.",
        type: "track-change",
      }),
      textLocator,
    ).execute();

    expect(result).toMatchObject({ success: true, commandId: "replace-1" });
    expect(env.cc.tag).toBe("stylistic:track-change:replace-1");
    expect(env.cc.title.startsWith(IDENTITY_TITLE_PREFIX)).toBe(true);

    const payload = JSON.parse(env.cc.title.slice(IDENTITY_TITLE_PREFIX.length));
    expect(payload).toEqual({
      suggestionId: "replace-1",
      version: "compound-v2",
      insertedSideRef: {
        kind: "content-control",
        role: "inserted-side",
        value: "stylistic:track-change:replace-1",
      },
      deletedSideRef: {
        kind: "anchor",
        role: "deleted-side",
        value: "texto original",
      },
      anchorRef: {
        kind: "anchor",
        role: "operational-anchor",
        value: "Contexto con texto original.",
      },
    });
  });

  it("returns a localized mutation patch for replace suggestions", async () => {
    const paragraphText = "Antes texto original y después.";
    const anchorRange = createRange({
      text: "texto original",
      paragraphText,
    });
    const contextRange = createRange({
      text: paragraphText,
      paragraphText,
      searchSequence: [[anchorRange]],
    });

    installWordContext({
      documentText: paragraphText,
      contextText: paragraphText,
      contextSearchSequence: [[contextRange]],
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "replace-patch-1",
        anchor: "texto original",
        suggestedText: "texto sugerido",
        context: paragraphText,
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({
      success: true,
      commandId: "replace-patch-1",
      mutationPatch: {
        suggestionId: "replace-patch-1",
        snapshotVersion: 1,
        originalText: paragraphText,
        updatedText: "Antes texto sugerido y después.",
        deltaLength: 0,
        affectedStart: 6,
        affectedEnd: 20,
      },
    });
  });

  it("keeps legacy anchor titles for comment-only suggestions", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "comment-1",
        type: "comment-only",
        suggestedText: undefined,
        anchor: "texto original",
      }),
      textLocator,
    ).execute();

    expect(result).toEqual({ success: true, commandId: "comment-1" });
    expect(env.cc.title).toBe("texto original");
  });
});
