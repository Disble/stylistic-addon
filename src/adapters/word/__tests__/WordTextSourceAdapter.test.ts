import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  installRejectingWord,
  installWordWithContext,
  makeParagraph,
} from "./WordAdapterTestHelper";
import { WordTextSourceAdapter } from "../WordTextSourceAdapter";

describe("WordTextSourceAdapter", () => {
  let adapter: WordTextSourceAdapter;
  let logSpy: ReturnType<typeof vi.spyOn>;
  const documentUuid = "11111111-1111-4111-8111-111111111111";

  beforeEach(() => {
    adapter = new WordTextSourceAdapter({
      getDocumentUuid: vi.fn().mockResolvedValue(documentUuid),
    } as never);
    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
  });

  afterEach(() => {
    logSpy.mockRestore();
    vi.unstubAllGlobals();
  });

  it("returns the active selection when it contains non-whitespace text", async () => {
    const selection = {
      load: vi.fn(),
      text: "Texto seleccionado",
      paragraphs: { items: [makeParagraph("Texto seleccionado")], load: vi.fn() },
    };
    const body = {
      load: vi.fn(),
      text: "Texto del documento",
      paragraphs: { items: [makeParagraph("Texto del documento")], load: vi.fn() },
    };
    const context = {
      document: { getSelection: vi.fn(() => selection), body },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.getTextToAnalyze()).resolves.toEqual({
      text: "Texto seleccionado",
      isSelection: true,
      documentUuid,
    });
  });

  it("propagates Word.run errors", async () => {
    installRejectingWord(new Error("Office host unavailable"));

    await expect(adapter.getTextToAnalyze()).rejects.toThrow("Office host unavailable");
  });
});
