import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "../WordAdapter";
import type { TextLocator } from "../WordTextLocatorContext.types";
import {
  installRejectingWord,
  installWordWithContext,
  makeOperationalWrapperTitle,
  makeOperationalWrapperTag,
  makeSuggestion,
} from "./WordAdapterActionTestHelper";

type LocatedRange = Awaited<ReturnType<NonNullable<TextLocator["locate"]>>>;

describe("WordAdapter.navigateToText", () => {
  let adapter: WordAdapter;

  beforeEach(() => {
    vi.clearAllMocks();
    adapter = new WordAdapter();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("selects the suggestion content control range when the real artifact exists", async () => {
    const select = vi.fn();
    const selectedRange = { select };
    const ccResult = {
      items: [
        {
          tag: makeOperationalWrapperTag("s-1"),
          title: makeOperationalWrapperTitle({
            deletedValue: "fragmento exacto",
            anchorValue: "Contexto con fragmento exacto.",
          }),
          load: vi.fn(),
          getRange: vi.fn(() => selectedRange),
        },
      ],
      load: vi.fn(),
    };
    const context = {
      document: {
        contentControls: {
          getByTag: vi.fn(() => ccResult),
        },
        body: {
          search: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(
      adapter.navigateToText(
        makeSuggestion({
          anchor: "fragmento exacto",
          context: "Contexto con fragmento exacto.",
        })
      )
    ).resolves.toEqual({ status: "navigated" });

    expect(context.document.contentControls.getByTag).toHaveBeenCalledWith(
      makeOperationalWrapperTag("s-1")
    );
    expect(select).toHaveBeenCalledOnce();
    expect(context.document.body.search).not.toHaveBeenCalled();
  });

  it("falls back to contextual anchor search when the content control is missing", async () => {
    const paragraphRange = {
      load: vi.fn(),
      text: "Contexto con fragmento exacto.",
      search: vi.fn(() => ({ items: [{ select: vi.fn() }], load: vi.fn() })),
    };
    const contextRange: {
      load: ReturnType<typeof vi.fn>;
      text: string;
      search: ReturnType<typeof vi.fn>;
      paragraphs: {
        getFirst: ReturnType<typeof vi.fn>;
      };
    } = {
      load: vi.fn(),
      text: "Contexto con fragmento exacto.",
      search: vi.fn(),
      paragraphs: {
        getFirst: vi.fn(() => ({ getRange: vi.fn(() => paragraphRange) })),
      },
    };
    const contextSearchResults = {
      items: [contextRange],
      load: vi.fn(),
    };
    const anchorSelect = vi.fn();
    const anchorSearchResults = {
      items: [{ select: anchorSelect }],
      load: vi.fn(),
    };
    const ccResult = {
      items: [],
      load: vi.fn(),
    };
    const body = {
      text: "Contexto con fragmento exacto.",
      load: vi.fn(),
      search: vi.fn().mockReturnValueOnce(contextSearchResults),
    };
    contextRange.search = vi.fn(() => anchorSearchResults);
    const context = {
      document: {
        contentControls: {
          getByTag: vi.fn(() => ccResult),
        },
        body,
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(
      adapter.navigateToText(
        makeSuggestion({
          anchor: "fragmento exacto",
          context: "Contexto con fragmento exacto.",
        })
      )
    ).resolves.toEqual({ status: "navigated" });

    expect(body.search).toHaveBeenCalledWith("Contexto con fragmento exacto.", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(contextRange.search).toHaveBeenCalledWith("fragmento exacto", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(anchorSelect).toHaveBeenCalledOnce();
  });

  it("selects the unique valid operational-wrapper CC and ignores stale drifted metadata", async () => {
    const staleSelect = vi.fn();
    const currentSelect = vi.fn();
    const ccResult = {
      items: [
        {
          tag: makeOperationalWrapperTag("chunk0-0"),
          title: makeOperationalWrapperTitle({
            suggestionId: "chunk0-0",
            insertedTag: "stylistic:track-change:chunk0-0",
            deletedValue: "anchor viejo",
            anchorValue: "Contexto viejo.",
          }),
          load: vi.fn(),
          getRange: vi.fn(() => ({ select: staleSelect })),
        },
        {
          tag: makeOperationalWrapperTag("chunk0-0"),
          title: makeOperationalWrapperTitle({
            suggestionId: "chunk0-0",
            insertedTag: "stylistic:track-change:chunk0-0",
            deletedValue: "fragmento actual",
            anchorValue: "Contexto con fragmento actual.",
          }),
          load: vi.fn(),
          getRange: vi.fn(() => ({ select: currentSelect })),
        },
      ],
      load: vi.fn(),
    };
    const context = {
      document: {
        contentControls: {
          getByTag: vi.fn(() => ccResult),
        },
        body: {
          search: vi.fn(),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(
      adapter.navigateToText(
        makeSuggestion({
          id: "chunk0-0",
          anchor: "fragmento actual",
          context: "Contexto con fragmento actual.",
        })
      )
    ).resolves.toEqual({ status: "navigated" });

    expect(currentSelect).toHaveBeenCalledOnce();
    expect(staleSelect).not.toHaveBeenCalled();
    expect(context.document.body.search).not.toHaveBeenCalled();
  });

  it("selects the first matching range when Word finds the text", async () => {
    const select = vi.fn();
    const results = {
      items: [{ select }],
      load: vi.fn(),
    };
    const context = {
      document: {
        body: {
          search: vi.fn(() => results),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.navigateToText("fragmento exacto")).resolves.toEqual({
      status: "navigated",
    });

    expect(context.document.body.search).toHaveBeenCalledWith("fragmento exacto", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(results.load).toHaveBeenCalledWith("items");
    expect(select).toHaveBeenCalledOnce();
    expect(context.sync).toHaveBeenCalledTimes(2);
  });

  it("delegates plain-string navigation to the injected text locator", async () => {
    const select = vi.fn();
    const locatedRange = { select } as unknown as LocatedRange;
    const locate = vi.fn(async () => locatedRange);
    const textLocator: TextLocator = { locate };
    const body = {
      search: vi.fn(),
      load: vi.fn(),
      text: "fragmento exacto",
    };
    const context = {
      document: { body },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);
    adapter = new WordAdapter(textLocator);

    await expect(adapter.navigateToText("fragmento exacto")).resolves.toEqual({
      status: "navigated",
    });

    expect(locate).toHaveBeenCalledOnce();
    expect(locate).toHaveBeenCalledWith({
      context,
      container: body,
      searchText: "fragmento exacto",
    });
    expect(select).toHaveBeenCalledOnce();
    expect(body.search).not.toHaveBeenCalled();
  });

  it("delegates suggestion fallback navigation to the injected text locator", async () => {
    const anchorSelect = vi.fn();
    const anchorRange = { select: anchorSelect } as unknown as LocatedRange;
    const contextRange = {
      load: vi.fn(),
      text: "Contexto con fragmento exacto.",
      paragraphs: {
        getFirst: vi.fn(() => ({
          getRange: vi.fn(() => ({ load: vi.fn(), text: "unused" })),
        })),
      },
    } as unknown as LocatedRange;
    const locate = vi
      .fn<NonNullable<TextLocator["locate"]>>()
      .mockResolvedValueOnce(contextRange)
      .mockResolvedValueOnce(anchorRange);
    const textLocator: TextLocator = { locate };
    const ccResult = {
      items: [],
      load: vi.fn(),
    };
    const body = {
      search: vi.fn(),
      load: vi.fn(),
      text: "Contexto con fragmento exacto.",
    };
    const context = {
      document: {
        contentControls: {
          getByTag: vi.fn(() => ccResult),
        },
        body,
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);
    adapter = new WordAdapter(textLocator);

    const suggestion = makeSuggestion({
      anchor: "fragmento exacto",
      context: "Contexto con fragmento exacto.",
    });

    await expect(adapter.navigateToText(suggestion)).resolves.toEqual({
      status: "navigated",
    });

    expect(locate).toHaveBeenCalledTimes(2);
    expect(locate).toHaveBeenNthCalledWith(1, {
      context,
      container: body,
      searchText: suggestion.context,
    });
    expect(locate).toHaveBeenNthCalledWith(2, {
      context,
      container: contextRange,
      searchText: suggestion.anchor,
    });
    expect(anchorSelect).toHaveBeenCalledOnce();
    expect(body.search).not.toHaveBeenCalled();
  });

  it("does nothing when the search returns no matches", async () => {
    const results = {
      items: [],
      load: vi.fn(),
    };
    const context = {
      document: {
        body: {
          load: vi.fn(),
          text: "",
          search: vi.fn(() => results),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.navigateToText("ausente")).resolves.toEqual({
      status: "not-found",
      reason: "plain-text-not-found",
    });

    expect(results.load).toHaveBeenCalledWith("items");
    expect(context.document.body.load).toHaveBeenCalledWith("text");
    expect(context.sync).toHaveBeenCalledTimes(3);
  });

  it("swallows Word host failures because navigation is best-effort", async () => {
    const run = installRejectingWord(new Error("Office host unavailable"));

    await expect(adapter.navigateToText("fragmento exacto")).resolves.toEqual({
      status: "failed",
      reason: "word-error",
    });

    expect(run).toHaveBeenCalledOnce();
  });
});
