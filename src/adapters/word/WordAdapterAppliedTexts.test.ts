import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "./WordAdapter";
import {
  installRejectingWord,
  installWordWithContext,
} from "./WordAdapterTestHelper";

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

  it("SR-DG-01: collects texts from both track-change and comment-only stylistic content controls", async () => {
    const rangeTC = { load: vi.fn(), text: "originalText1" };
    const rangeCO = { load: vi.fn(), text: "originalText2" };
    const rangeOther = { load: vi.fn(), text: "should not appear" };

    const context = {
      document: {
        contentControls: {
          items: [
            {
              tag: "stylistic:track-change:s1",
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
    expect(rangeTC.load).toHaveBeenCalledWith("text");
    expect(rangeCO.load).toHaveBeenCalledWith("text");
    expect(context.sync).toHaveBeenCalledTimes(2);
  });

  it("deduplicates texts when multiple stylistic content controls span the same text", async () => {
    const range1 = { load: vi.fn(), text: "texto duplicado" };
    const range2 = { load: vi.fn(), text: "texto duplicado" };
    const range3 = { load: vi.fn(), text: "texto único" };

    const context = {
      document: {
        contentControls: {
          items: [
            { tag: "stylistic:track-change:s1", getRange: vi.fn(() => range1) },
            { tag: "stylistic:track-change:s2", getRange: vi.fn(() => range2) },
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

  it("propagates Word.run errors", async () => {
    installRejectingWord(new Error("Tracked changes unavailable"));

    await expect(adapter.getAppliedOriginalTexts()).rejects.toThrow(
      "Tracked changes unavailable",
    );
  });
});
