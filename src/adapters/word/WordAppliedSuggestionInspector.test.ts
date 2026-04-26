import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { installWordWithContext } from "./WordAdapterTestHelper";
import { WordAppliedSuggestionInspector } from "./WordAppliedSuggestionInspector";

describe("WordAppliedSuggestionInspector", () => {
  let inspector: WordAppliedSuggestionInspector;
  let logSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    inspector = new WordAppliedSuggestionInspector();
    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
  });

  afterEach(() => {
    logSpy.mockRestore();
    vi.unstubAllGlobals();
  });

  it("returns an empty set when there are no stylistic content controls", async () => {
    const context = {
      document: {
        contentControls: {
          items: [{ tag: "other-cc", getRange: vi.fn() }],
          load: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(inspector.getAppliedOriginalTexts()).resolves.toEqual(new Set());
  });
});
