import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "../WordAdapter";
import {
  installRejectingWord,
  installWordWithContext,
  makeParagraph,
} from "../WordAdapterTestHelper";

describe("WordAdapter.getTextToAnalyze", () => {
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

  it("returns the active selection when it contains non-whitespace text", async () => {
    const selection = {
      load: vi.fn(),
      text: "Texto seleccionado",
      paragraphs: {
        items: [makeParagraph("Texto seleccionado")],
        load: vi.fn(),
      },
    };
    const body = {
      load: vi.fn(),
      text: "Texto del documento",
      paragraphs: {
        items: [makeParagraph("Texto del documento")],
        load: vi.fn(),
      },
    };
    const context = {
      document: {
        getSelection: vi.fn(() => selection),
        body,
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    const run = installWordWithContext(context);

    await expect(adapter.getTextToAnalyze()).resolves.toEqual({
      text: "Texto seleccionado",
      isSelection: true,
    });

    expect(run).toHaveBeenCalledOnce();
    expect(selection.load).toHaveBeenCalledWith("text,paragraphs/items");
    expect(selection.paragraphs.load).toHaveBeenCalledWith(
      "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent"
    );
    expect(body.load).not.toHaveBeenCalled();
    expect(body.paragraphs.load).not.toHaveBeenCalled();
    expect(context.sync).toHaveBeenCalledTimes(1);
  });

  it("falls back to the full document body when selection is empty or whitespace", async () => {
    const selection = {
      load: vi.fn(),
      text: "   \n  ",
      paragraphs: {
        items: [],
        load: vi.fn(),
      },
    };
    const body = {
      load: vi.fn(),
      text: "Texto completo del documento",
      paragraphs: {
        items: [makeParagraph("Texto completo del documento")],
        load: vi.fn(),
      },
    };
    const context = {
      document: {
        getSelection: vi.fn(() => selection),
        body,
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.getTextToAnalyze()).resolves.toEqual({
      text: "Texto completo del documento",
      isSelection: false,
    });

    expect(body.paragraphs.load).toHaveBeenCalledWith(
      "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent"
    );
    expect(body.load).toHaveBeenCalledWith("paragraphs/items");
    expect(context.sync).toHaveBeenCalledTimes(2);
  });

  it("falls back to the full document when the cursor is collapsed inside a paragraph", async () => {
    const selection = {
      load: vi.fn(),
      text: "   ",
      paragraphs: {
        items: [makeParagraph("Párrafo actual")],
        load: vi.fn(),
      },
    };
    const body = {
      load: vi.fn(),
      text: "Título Párrafo actual Segundo párrafo con sangría.",
      paragraphs: {
        items: [
          makeParagraph("Título", { styleBuiltIn: "Title" }),
          makeParagraph("Párrafo actual"),
          makeParagraph("Segundo párrafo con sangría.", { firstLineIndent: 18 }),
        ],
        load: vi.fn(),
      },
    };
    const context = {
      document: {
        getSelection: vi.fn(() => selection),
        body,
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.getTextToAnalyze()).resolves.toEqual({
      text: "Título\n\nPárrafo actual\n\n\tSegundo párrafo con sangría.",
      isSelection: false,
    });
  });

  it("keeps title paragraphs visually separated from body paragraphs", async () => {
    const selection = {
      load: vi.fn(),
      text: "",
      paragraphs: {
        items: [],
        load: vi.fn(),
      },
    };
    const body = {
      load: vi.fn(),
      text: "Capítulo 1 Primer párrafo del capítulo.",
      paragraphs: {
        items: [
          makeParagraph("Capítulo 1", { styleBuiltIn: "Title" }),
          makeParagraph("Primer párrafo del capítulo."),
        ],
        load: vi.fn(),
      },
    };
    const context = {
      document: {
        getSelection: vi.fn(() => selection),
        body,
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.getTextToAnalyze()).resolves.toEqual({
      text: "Capítulo 1\n\nPrimer párrafo del capítulo.",
      isSelection: false,
    });
  });

  it("reflects paragraph indentation metadata with an explicit tab prefix", async () => {
    const selection = {
      load: vi.fn(),
      text: "",
      paragraphs: {
        items: [],
        load: vi.fn(),
      },
    };
    const body = {
      load: vi.fn(),
      text: "Primer párrafo. Segundo párrafo con sangría visual.",
      paragraphs: {
        items: [
          makeParagraph("Primer párrafo."),
          makeParagraph("Segundo párrafo con sangría visual.", {
            firstLineIndent: 18,
          }),
        ],
        load: vi.fn(),
      },
    };
    const context = {
      document: {
        getSelection: vi.fn(() => selection),
        body,
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.getTextToAnalyze()).resolves.toEqual({
      text: "Primer párrafo.\n\n\tSegundo párrafo con sangría visual.",
      isSelection: false,
    });
  });

  it("returns an empty body text unchanged when the document is empty", async () => {
    const selection = {
      load: vi.fn(),
      text: "",
      paragraphs: {
        items: [],
        load: vi.fn(),
      },
    };
    const body = {
      load: vi.fn(),
      text: "",
      paragraphs: {
        items: [],
        load: vi.fn(),
      },
    };
    const context = {
      document: {
        getSelection: vi.fn(() => selection),
        body,
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.getTextToAnalyze()).resolves.toEqual({
      text: "",
      isSelection: false,
    });
  });

  it("propagates Word.run errors", async () => {
    const error = new Error("Office host unavailable");
    const run = installRejectingWord(error);

    await expect(adapter.getTextToAnalyze()).rejects.toThrow("Office host unavailable");

    expect(run).toHaveBeenCalledOnce();
  });
});
