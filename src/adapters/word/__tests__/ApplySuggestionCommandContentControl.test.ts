import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ApplySuggestionCommand } from "../ApplySuggestionCommand";
import { WordTextLocatorAdapter } from "../WordTextLocatorAdapter";
import {
  createRange,
  installWordContext,
  makeSuggestion,
} from "../ApplySuggestionCommandTestHelper";
import type { ParentCC } from "../ApplySuggestionCommandTestHelper.types";

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
      textLocator
    ).execute();

    expect(result).toMatchObject({ success: true, commandId: "s1" });
    expect(env.anchorRange.insertComment).toHaveBeenCalledWith("[Estilo]\nMejora la claridad");
    expect(env.anchorRange.insertContentControl).toHaveBeenCalledOnce();
  });

  it("aborts before mutation when an existing stylistic content control covers the anchor", async () => {
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

    const result = await new ApplySuggestionCommand(makeSuggestion(), textLocator).execute();

    expect(result).toMatchObject({ success: false, commandId: "s1" });
    expect(coveredParentCC.delete).not.toHaveBeenCalled();
    expect(env.context.document.body.search).toHaveBeenCalledTimes(1);
    expect(freshAnchor.insertText).not.toHaveBeenCalled();
  });

  it("aborts before mutation when a chunk wrapper covers the anchor", async () => {
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

    const result = await new ApplySuggestionCommand(makeSuggestion(), textLocator).execute();

    expect(result).toMatchObject({ success: false, commandId: "s1" });
    expect(coveredParentCC.delete).not.toHaveBeenCalled();
    expect(freshAnchor.insertText).not.toHaveBeenCalled();
  });

  it("persists operational-wrapper replace metadata for replace suggestions", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "replace-1",
        anchor: "texto original",
        suggestedText: "texto sugerido",
        context: "Contexto con texto original.",
        type: "track-change",
      }),
      textLocator
    ).execute();

    expect(result).toMatchObject({ success: true, commandId: "replace-1" });
    expect(env.operationalWrapper.tag).toBe("stylistic-operational-wrapper:replace-1");
    expect(env.operationalWrapper.title.startsWith(IDENTITY_TITLE_PREFIX)).toBe(true);
    expect(env.cc.tag).toBe("stylistic:track-change:replace-1");
    expect(env.cc.title).toBe("texto original");

    const payload = JSON.parse(env.operationalWrapper.title.slice(IDENTITY_TITLE_PREFIX.length));
    expect(payload).toEqual({
      suggestionId: "replace-1",
      version: "operational-wrapper-v1",
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
      groupId: "replace-1",
      groupIndex: 0,
      groupSize: 1,
    });
  });

  it("reuses an existing matching operational wrapper instead of creating a duplicate", async () => {
    const env = installWordContext({
      anchorRangeParentCC: {
        tag: "stylistic-operational-wrapper:replace-existing",
        isNullObject: false,
        load: vi.fn(),
        delete: vi.fn(),
      },
    });

    env.anchorRange.parentContentControlOrNullObject.tag =
      "stylistic-operational-wrapper:replace-existing";
    env.anchorRange.parentContentControlOrNullObject.isNullObject = false;
    env.anchorRange.parentContentControlOrNullObject.title = env.operationalWrapper.title;
    env.anchorRange.parentContentControlOrNullObject.getRange = env.operationalWrapper.getRange;
    env.operationalWrapper.tag = "stylistic-operational-wrapper:replace-existing";
    env.operationalWrapper.title = `${IDENTITY_TITLE_PREFIX}${JSON.stringify({
      suggestionId: "replace-existing",
      version: "operational-wrapper-v1",
      insertedSideRef: {
        kind: "content-control",
        role: "inserted-side",
        value: "stylistic:track-change:replace-existing",
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
      groupId: "replace-existing",
      groupIndex: 0,
      groupSize: 1,
    })}`;
    env.anchorRange.parentContentControlOrNullObject.title = env.operationalWrapper.title;

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "replace-existing",
        anchor: "texto original",
        suggestedText: "texto sugerido",
        context: "Contexto con texto original.",
        type: "track-change",
      }),
      textLocator
    ).execute();

    expect(result).toMatchObject({ success: true, commandId: "replace-existing" });
    expect(env.anchorRange.insertContentControl).not.toHaveBeenCalled();
    expect(env.operationalWrapper.getRange).toHaveBeenCalled();
  });

  it("aborts before mutation when the anchor is covered by a duplicated operational wrapper", async () => {
    const coveredParentCC: ParentCC = {
      tag: "stylistic-operational-wrapper:replace-dup",
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

    installWordContext({
      contextSearchSequence: [[coveredContext]],
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "replace-dup",
        anchor: "texto original",
        suggestedText: "texto sugerido",
        context: "Contexto con texto original.",
      }),
      textLocator
    ).execute();

    expect(result).toEqual({
      success: false,
      commandId: "replace-dup",
      error: "Anchor cubierto por un Content Control existente",
    });
    expect(coveredAnchor.insertText).not.toHaveBeenCalled();
  });

  it("returns a localized mutation patch for replace suggestions", async () => {
    const paragraphText = "Antes texto original y después.";
    installWordContext({
      documentText: paragraphText,
      contextText: paragraphText,
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "replace-patch-1",
        anchor: "texto original",
        suggestedText: "texto sugerido",
        context: paragraphText,
      }),
      textLocator
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

  it("re-locates the current inserted side before annotating a replace suggestion", async () => {
    const insertedHybridRange = createRange({
      text: "texto sugerido texto original",
      reviewedCurrentText: "texto sugerido",
      reviewedOriginalText: "texto original",
      paragraphText: "Antes texto sugerido texto original y después.",
    });
    const isolatedInsertedRange = createRange({
      text: "texto sugerido",
      reviewedCurrentText: "texto sugerido",
      reviewedOriginalText: "",
      paragraphText: "Antes texto sugerido texto original y después.",
    });
    insertedHybridRange.search = vi
      .fn()
      .mockReturnValueOnce({ items: [isolatedInsertedRange], load: vi.fn() });

    installWordContext({
      insertedRange: {
        text: insertedHybridRange.text,
        reviewedCurrentText: "texto sugerido",
        reviewedOriginalText: "texto original",
        searchSequence: [[isolatedInsertedRange]],
        paragraphText: "Antes texto sugerido texto original y después.",
      },
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "replace-rerange-1",
        anchor: "texto original",
        suggestedText: "texto sugerido",
        context: "Contexto con texto original.",
        type: "track-change",
      }),
      textLocator
    ).execute();

    expect(result).toMatchObject({ success: true, commandId: "replace-rerange-1" });
    expect(isolatedInsertedRange.insertComment).toHaveBeenCalledWith(
      "[Estilo]\nMejora la claridad"
    );
    expect(isolatedInsertedRange.insertContentControl).toHaveBeenCalledOnce();
  });

  it("annotates delete-only suggestions on the operational wrapper instead of an inserted side", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "delete-only-1",
        anchor: "texto original",
        suggestedText: "",
        context: "Contexto con texto original.",
        type: "track-change",
      }),
      textLocator
    ).execute();

    expect(result).toMatchObject({
      success: true,
      commandId: "delete-only-1",
      mutationPatch: {
        suggestionId: "delete-only-1",
        snapshotVersion: 1,
        originalText: "Contexto con texto original.",
        updatedText: "Contexto con .",
        deltaLength: -14,
        affectedStart: 13,
        affectedEnd: 27,
      },
    });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith("", "Replace");
    expect(env.operationalWrapperRange.insertComment).toHaveBeenCalledWith(
      "[Estilo]\nMejora la claridad"
    );
    expect(env.operationalWrapperRange.insertContentControl).toHaveBeenCalledOnce();
    expect(env.insertedRange.insertComment).not.toHaveBeenCalled();
    expect(env.insertedRange.insertContentControl).not.toHaveBeenCalled();
  });

  it("applies markdown typography as native formatting without inserting literal asterisks", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "format-italic-1",
        anchor: "texto original",
        suggestedText: "*texto original*",
        context: "Contexto con texto original.",
        type: "track-change",
      }),
      textLocator
    ).execute();

    expect(result).toEqual({ success: true, commandId: "format-italic-1" });
    expect(env.anchorRange.font.italic).toBe(true);
    expect(env.anchorRange.font.bold).toBe(false);
    expect(env.anchorRange.insertText).not.toHaveBeenCalled();
    expect(env.anchorRange.insertComment).toHaveBeenCalledWith("[Estilo]\nMejora la claridad");
    expect(env.anchorRange.insertContentControl).toHaveBeenCalledTimes(2);
  });

  it("fails safely when it cannot isolate the current inserted side for a replace suggestion", async () => {
    installWordContext({
      insertedRange: {
        text: "texto sugerido texto original",
        reviewedCurrentText: "texto sugerido",
        reviewedOriginalText: "texto original",
        searchSequence: [[]],
        paragraphText: "Antes texto sugerido texto original y después.",
      },
    });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "replace-rerange-fail-1",
        anchor: "texto original",
        suggestedText: "texto sugerido",
        context: "Contexto con texto original.",
        type: "track-change",
      }),
      textLocator
    ).execute();

    expect(result).toEqual({
      success: false,
      commandId: "replace-rerange-fail-1",
      error: "No se pudo aislar el texto insertado de la sugerencia",
    });
  });

  it("keeps legacy anchor titles for comment-only suggestions", async () => {
    const env = installWordContext({ useOperationalWrapper: false });

    const result = await new ApplySuggestionCommand(
      makeSuggestion({
        id: "comment-1",
        type: "comment-only",
        suggestedText: undefined,
        anchor: "texto original",
      }),
      textLocator
    ).execute();

    expect(result).toEqual({ success: true, commandId: "comment-1" });
    expect(env.cc.title).toBe("texto original");
  });
});
