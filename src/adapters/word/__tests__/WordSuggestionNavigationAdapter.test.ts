import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { TextLocator } from "../WordTextLocatorContext.types";
import { installWordWithContext, makeSuggestion } from "../WordAdapterActionTestHelper";
import { WordSuggestionNavigationAdapter } from "../WordSuggestionNavigationAdapter";

describe("WordSuggestionNavigationAdapter", () => {
  let locate: ReturnType<typeof vi.fn<TextLocator["locate"]>>;
  let adapter: WordSuggestionNavigationAdapter;

  /** Returns a minimal Word range accepted by the text locator contract in tests. */
  function makeLocatedRange(): Word.Range {
    return { select: vi.fn() } as unknown as Word.Range;
  }

  beforeEach(() => {
    locate = vi.fn<TextLocator["locate"]>();
    const textLocator: TextLocator = { locate };
    adapter = new WordSuggestionNavigationAdapter(textLocator);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("delegates plain-string navigation to the injected text locator", async () => {
    const range = makeLocatedRange();
    locate.mockResolvedValue(range);
    const body = { search: vi.fn(), load: vi.fn(), text: "fragmento exacto" };
    const context = {
      document: { body },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.navigateToText("fragmento exacto")).resolves.toEqual({
      status: "navigated",
    });
    expect(locate).toHaveBeenCalledOnce();
  });

  it("does not globally search the anchor when contextual fallback cannot find context", async () => {
    locate.mockResolvedValue(null);
    const ccResult = { items: [], load: vi.fn() };
    const body = { search: vi.fn(), load: vi.fn(), text: "donde aparece en TOC" };
    const context = {
      document: {
        contentControls: { getByTag: vi.fn(() => ccResult) },
        body,
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const suggestion = makeSuggestion({
      anchor: "donde",
      context: "Contexto real ausente con donde.",
    });

    installWordWithContext(context);

    await expect(adapter.navigateToText(suggestion)).resolves.toEqual({
      status: "not-found",
      reason: "context-not-found",
    });

    expect(locate).toHaveBeenCalledTimes(1);
    expect(locate).toHaveBeenCalledWith({
      context,
      container: body,
      searchText: suggestion.context,
    });
    expect(body.search).not.toHaveBeenCalled();
  });
});
