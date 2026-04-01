import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "./WordAdapter";
import {
  installRejectingWord,
  installWordWithContext,
} from "./WordAdapterTestHelper";

describe("WordAdapter.navigateToText", () => {
  let adapter: WordAdapter;

  beforeEach(() => {
    vi.clearAllMocks();
    adapter = new WordAdapter();
  });

  afterEach(() => {
    vi.unstubAllGlobals();
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

    await expect(adapter.navigateToText("fragmento exacto")).resolves.toBeUndefined();

    expect(context.document.body.search).toHaveBeenCalledWith("fragmento exacto", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(results.load).toHaveBeenCalledWith("items");
    expect(select).toHaveBeenCalledOnce();
    expect(context.sync).toHaveBeenCalledTimes(2);
  });

  it("does nothing when the search returns no matches", async () => {
    const results = {
      items: [],
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

    await expect(adapter.navigateToText("ausente")).resolves.toBeUndefined();

    expect(results.load).toHaveBeenCalledWith("items");
    expect(context.sync).toHaveBeenCalledTimes(1);
  });

  it("swallows Word host failures because navigation is best-effort", async () => {
    const run = installRejectingWord(new Error("Office host unavailable"));

    await expect(adapter.navigateToText("fragmento exacto")).resolves.toBeUndefined();

    expect(run).toHaveBeenCalledOnce();
  });
});
