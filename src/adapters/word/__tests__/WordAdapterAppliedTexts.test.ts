import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "../WordAdapter";
import { installRejectingWord, installWordWithContext } from "../WordAdapterTestHelper";

describe("WordAdapter.getAppliedOriginalTexts", () => {
  let adapter: WordAdapter;
  let logSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.clearAllMocks();
    adapter = new WordAdapter();
    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
  });

  afterEach(() => {
    logSpy.mockRestore();
    vi.unstubAllGlobals();
  });

  it("SR-DG-02: returns an empty set when there are no stylistic content controls", async () => {
    const context = {
      document: {
        contentControls: {
          items: [
            { tag: "other-cc", getRange: vi.fn() },
            { tag: "chunk0-0", getRange: vi.fn() },
          ],
          load: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    const result = await adapter.getAppliedOriginalTexts();

    expect(result).toEqual(new Set());
    expect(context.document.contentControls.items[0].getRange).not.toHaveBeenCalled();
    expect(context.document.contentControls.items[1].getRange).not.toHaveBeenCalled();
    expect(context.sync).toHaveBeenCalledTimes(1);
  });

  it("SR-DG-01: collects comment-only titles and only operational-wrapper track-change refs", async () => {
    const rangeTC = { load: vi.fn(), text: "originalText1" };
    const rangeCO = { load: vi.fn(), text: "originalText2" };
    const rangeOther = { load: vi.fn(), text: "should not appear" };

    const context = {
      document: {
        contentControls: {
          items: [
            {
              tag: "stylistic:track-change:s1",
              title:
                'stylistic-meta-v2:{"suggestionId":"s1","version":"operational-wrapper-v1","insertedSideRef":{"kind":"content-control","role":"inserted-side","value":"stylistic:track-change:s1"},"deletedSideRef":{"kind":"anchor","role":"deleted-side","value":"originalText1"},"anchorRef":{"kind":"anchor","role":"operational-anchor","value":"Contexto con texto original."},"groupId":"s1","groupIndex":0,"groupSize":1}',
              getRange: vi.fn(() => rangeTC),
            },
            {
              tag: "stylistic:comment-only:s2",
              getRange: vi.fn(() => rangeCO),
            },
            { tag: "other-cc", getRange: vi.fn(() => rangeOther) },
          ],
          load: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    const result = await adapter.getAppliedOriginalTexts();

    expect(result).toEqual(new Set(["originalText1", "originalText2"]));
    expect(rangeTC.load).not.toHaveBeenCalledWith("text");
    expect(rangeCO.load).toHaveBeenCalledWith("text");
    expect(context.sync).toHaveBeenCalledTimes(2);
  });

  it("deduplicates texts when multiple Stylistic content controls map to the same original text", async () => {
    const range1 = { load: vi.fn(), text: "texto duplicado" };
    const range2 = { load: vi.fn(), text: "texto duplicado" };
    const range3 = { load: vi.fn(), text: "texto único" };

    const context = {
      document: {
        contentControls: {
          items: [
            {
              tag: "stylistic:track-change:s1",
              title:
                'stylistic-meta-v2:{"suggestionId":"s1","version":"operational-wrapper-v1","insertedSideRef":{"kind":"content-control","role":"inserted-side","value":"stylistic:track-change:s1"},"deletedSideRef":{"kind":"anchor","role":"deleted-side","value":"texto duplicado"},"anchorRef":{"kind":"anchor","role":"operational-anchor","value":"Contexto 1"},"groupId":"s1","groupIndex":0,"groupSize":1}',
              getRange: vi.fn(() => range1),
            },
            {
              tag: "stylistic:track-change:s2",
              title:
                'stylistic-meta-v2:{"suggestionId":"s2","version":"operational-wrapper-v1","insertedSideRef":{"kind":"content-control","role":"inserted-side","value":"stylistic:track-change:s2"},"deletedSideRef":{"kind":"anchor","role":"deleted-side","value":"texto duplicado"},"anchorRef":{"kind":"anchor","role":"operational-anchor","value":"Contexto 2"},"groupId":"s2","groupIndex":0,"groupSize":1}',
              getRange: vi.fn(() => range2),
            },
            { tag: "stylistic:comment-only:s3", getRange: vi.fn(() => range3) },
          ],
          load: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    const result = await adapter.getAppliedOriginalTexts();

    expect(result).toEqual(new Set(["texto duplicado", "texto único"]));
    expect(result.size).toBe(2);
  });

  it("prefers persisted original text metadata from ContentControl.title over mutated visible range text", async () => {
    const mutatedReplaceRange = { load: vi.fn(), text: "parece que es" };
    const mutatedCommentOnlyRange = { load: vi.fn(), text: "tomar" };

    const context = {
      document: {
        contentControls: {
          items: [
            {
              tag: "stylistic:track-change:s1",
              title:
                'stylistic-meta-v2:{"suggestionId":"s1","version":"operational-wrapper-v1","insertedSideRef":{"kind":"content-control","role":"inserted-side","value":"stylistic:track-change:s1"},"deletedSideRef":{"kind":"anchor","role":"deleted-side","value":"parece que es la"},"anchorRef":{"kind":"anchor","role":"operational-anchor","value":"Contexto con parece que es la."},"groupId":"s1","groupIndex":0,"groupSize":1}',
              getRange: vi.fn(() => mutatedReplaceRange),
            },
            {
              tag: "stylistic:comment-only:s2",
              title: "se cometen",
              getRange: vi.fn(() => mutatedCommentOnlyRange),
            },
          ],
          load: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    const result = await adapter.getAppliedOriginalTexts();

    expect(result).toEqual(new Set(["parece que es la", "se cometen"]));
    expect(mutatedReplaceRange.load).not.toHaveBeenCalledWith("text");
    expect(mutatedCommentOnlyRange.load).not.toHaveBeenCalledWith("text");
  });

  it("ignores track-change content controls without operational-wrapper metadata", async () => {
    const legacyRange = { load: vi.fn(), text: "texto legado" };

    const context = {
      document: {
        contentControls: {
          items: [
            {
              tag: "stylistic:track-change:legacy-s1",
              title: "",
              getRange: vi.fn(() => legacyRange),
            },
          ],
          load: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    const result = await adapter.getAppliedOriginalTexts();

    expect(result).toEqual(new Set());
    expect(legacyRange.load).not.toHaveBeenCalledWith("text");
  });

  it("prefers deleted-side refs from operational-wrapper metadata when collecting already applied originals", async () => {
    const context = {
      document: {
        contentControls: {
          items: [
            {
              tag: "stylistic:track-change:s1",
              title:
                'stylistic-meta-v2:{"suggestionId":"s1","version":"operational-wrapper-v1","insertedSideRef":{"kind":"content-control","role":"inserted-side","value":"stylistic:track-change:s1"},"deletedSideRef":{"kind":"anchor","role":"deleted-side","value":"texto original v2"},"anchorRef":{"kind":"anchor","role":"operational-anchor","value":"Contexto con texto original v2."},"groupId":"s1","groupIndex":0,"groupSize":1}',
              getRange: vi.fn(),
            },
          ],
          load: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    const result = await adapter.getAppliedOriginalTexts();

    expect(result).toEqual(new Set(["texto original v2"]));
    expect(context.document.contentControls.items[0].getRange).not.toHaveBeenCalled();
  });

  it("propagates Word.run errors", async () => {
    installRejectingWord(new Error("Tracked changes unavailable"));

    await expect(adapter.getAppliedOriginalTexts()).rejects.toThrow("Tracked changes unavailable");
  });
});
