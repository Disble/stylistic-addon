import { describe, expect, it, vi } from "vitest";
import { WordTextLocatorAdapter } from "../WordTextLocatorAdapter";

type MockRange = { id: string };

type SearchOutcome = { kind: "results"; items: MockRange[] } | { kind: "invalid" };

/** Creates a minimal fake RangeCollection for locator tests. */
function createRangeCollection(items: MockRange[]) {
  return {
    items,
    load: vi.fn(),
  };
}

/** Creates a minimal Word-like search container with sequenced outcomes. */
function createSearchContainer(text: string, outcomes: SearchOutcome[]) {
  return {
    text,
    load: vi.fn(),
    search: vi.fn((_searchText: string, _options: Record<string, boolean>) => {
      const next = outcomes.shift();
      if (!next) {
        return createRangeCollection([]);
      }

      if (next.kind === "invalid") {
        throw new Error("SearchStringInvalidOrTooLong");
      }

      return createRangeCollection(next.items);
    }),
  };
}

/** Creates a minimal request context mock for locator tests. */
function createRequestContext() {
  return {
    sync: vi.fn(async () => undefined),
  } as unknown as Word.RequestContext;
}

describe("WordTextLocatorAdapter", () => {
  it("returns the first exact-match result when Word finds it directly", async () => {
    const exactRange: MockRange = { id: "exact" };
    const adapter = new WordTextLocatorAdapter();
    const container = createSearchContainer("Contexto con texto original.", [
      { kind: "results", items: [exactRange] },
    ]);

    const result = await adapter.locate({
      context: createRequestContext(),
      container: container as never,
      searchText: "texto original",
    });

    expect(result).toBe(exactRange);
    expect(container.search).toHaveBeenCalledTimes(1);
    expect(container.search).toHaveBeenCalledWith("texto original", {
      matchCase: true,
      matchWholeWord: false,
    });
  });

  it("skips the exact search when the text exceeds Word's 256-char limit", async () => {
    const longSearchText = `Prefijo ${"x".repeat(270)}`;
    const relaxedRange: MockRange = { id: "relaxed" };
    const adapter = new WordTextLocatorAdapter();
    const container = createSearchContainer(longSearchText, [
      { kind: "results", items: [relaxedRange] },
    ]);

    const result = await adapter.locate({
      context: createRequestContext(),
      container: container as never,
      searchText: longSearchText,
    });

    expect(result).toBe(relaxedRange);
    expect(container.search).toHaveBeenCalledTimes(1);
    expect(container.search).toHaveBeenCalledWith(longSearchText, {
      matchCase: true,
      matchWholeWord: false,
      ignorePunct: true,
      ignoreSpace: true,
    });
  });

  it("falls back to a whitespace-insensitive slice when Word search returns no results", async () => {
    const fallbackRange: MockRange = { id: "fallback" };
    const adapter = new WordTextLocatorAdapter();
    const container = createSearchContainer("Contexto con texto\n\noriginal.", [
      { kind: "results", items: [] },
      { kind: "results", items: [] },
      { kind: "results", items: [fallbackRange] },
    ]);

    const result = await adapter.locate({
      context: createRequestContext(),
      container: container as never,
      searchText: "texto original",
    });

    expect(result).toBe(fallbackRange);
    expect(container.load).toHaveBeenCalledWith("text");
    expect(container.search).toHaveBeenNthCalledWith(3, "texto\n\noriginal", {
      matchCase: true,
      matchWholeWord: false,
    });
  });

  it("retries from the first alphanumeric offset when the fallback candidate is invalid", async () => {
    const retriedRange: MockRange = { id: "retried" };
    const adapter = new WordTextLocatorAdapter();
    const searchText = "—¿Sabes quién es la tercera? " + "A".repeat(260);
    const containerText = `${searchText} final`;
    const container = createSearchContainer(containerText, [
      { kind: "invalid" },
      { kind: "invalid" },
      { kind: "results", items: [retriedRange] },
    ]);

    const result = await adapter.locate({
      context: createRequestContext(),
      container: container as never,
      searchText,
    });

    expect(result).toBe(retriedRange);
    expect(container.search).toHaveBeenCalledTimes(3);

    const retriedCandidate = container.search.mock.calls[2]?.[0];
    expect(typeof retriedCandidate).toBe("string");
    expect(retriedCandidate.startsWith("—")).toBe(false);
    expect(retriedCandidate.startsWith("¿")).toBe(false);
  });
});
