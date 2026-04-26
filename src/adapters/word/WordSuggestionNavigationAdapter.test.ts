import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { TextLocator } from "./WordTextLocatorContext";
import { installWordWithContext, makeSuggestion } from "./WordAdapterActionTestHelper";
import { WordSuggestionNavigationAdapter } from "./WordSuggestionNavigationAdapter";

describe("WordSuggestionNavigationAdapter", () => {
  let locate: ReturnType<typeof vi.fn>;
  let adapter: WordSuggestionNavigationAdapter;

  beforeEach(() => {
    locate = vi.fn();
    adapter = new WordSuggestionNavigationAdapter({ locate } as TextLocator);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("delegates plain-string navigation to the injected text locator", async () => {
    const select = vi.fn();
    locate.mockResolvedValue({ select });
    const body = { search: vi.fn(), load: vi.fn(), text: "fragmento exacto" };
    const context = {
      document: { body },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.navigateToText("fragmento exacto")).resolves.toBeUndefined();
    expect(locate).toHaveBeenCalledOnce();
  });

  it("builds the canonical stylistic tag for a suggestion", () => {
    expect(adapter.buildSuggestionTag(makeSuggestion({ id: "s-1" }))).toBe(
      "stylistic:track-change:s-1",
    );
  });
});
