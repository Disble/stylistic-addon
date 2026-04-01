import type { CommandResult, Suggestion } from "../../domain/types";

const commandMocks = vi.hoisted(() => ({
  constructor: vi.fn<(suggestion: Suggestion) => void>(),
  execute: vi.fn<(suggestion: Suggestion) => Promise<CommandResult>>(),
}));

const cleanupMocks = vi.hoisted(() => ({
  cleanupResolvedComments: vi.fn(),
}));

vi.mock("./ApplySuggestionCommand", () => ({
  ApplySuggestionCommand: class {
    private readonly suggestion: Suggestion;

    constructor(suggestion: Suggestion) {
      this.suggestion = suggestion;
      commandMocks.constructor(suggestion);
    }

    execute() {
      return commandMocks.execute(this.suggestion);
    }
  },
}));

vi.mock("./cleanup/CommentCleanup", () => ({
  cleanupResolvedComments: cleanupMocks.cleanupResolvedComments,
  OVERLAPPING_RELATIONS: [
    "Equal",
    "Contains",
    "ContainsStart",
    "ContainsEnd",
    "Inside",
    "InsideStart",
    "InsideEnd",
    "OverlapsBefore",
    "OverlapsAfter",
  ],
  COMMENT_ONLY_TAG_PREFIX: "stylistic:comment-only:",
}));

import { WordAdapter } from "./WordAdapter";

type WordRunCallback<T> = (context: any) => Promise<T> | T;

function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  const anchor = overrides.anchor ?? "texto original";
  return {
    id: "s-1",
    context: overrides.context ?? `Contexto con ${anchor}.`,
    anchor,
    suggestedText: "texto sugerido",
    justification: "Mas claro",
    category: "Claridad",
    severity: "medium",
    type: "track-change",
    ...overrides,
  };
}

function installWordWithContext(context: any) {
  const run = vi.fn(async <T>(callback: WordRunCallback<T>) =>
    callback(context),
  );
  (globalThis as any).Word = { run };
  return run;
}

function installRejectingWord(error: Error) {
  const run = vi.fn().mockRejectedValue(error);
  (globalThis as any).Word = { run };
  return run;
}

function makeParagraph(
  text: string,
  overrides: Partial<{
    styleBuiltIn: string;
    firstLineIndent: number;
    leftIndent: number;
  }> = {},
) {
  return {
    text,
    styleBuiltIn: "Normal",
    firstLineIndent: 0,
    leftIndent: 0,
    ...overrides,
  };
}

describe("WordAdapter", () => {
  let adapter: WordAdapter;
  let logSpy: ReturnType<typeof vi.spyOn>;
  let warnSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.clearAllMocks();
    adapter = new WordAdapter();
    cleanupMocks.cleanupResolvedComments.mockResolvedValue({
      deleted: 0,
      kept: 0,
    });

    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
    warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
  });

  afterEach(() => {
    logSpy.mockRestore();
    warnSpy.mockRestore();
    delete (globalThis as any).Word;
  });

  describe("getTextToAnalyze", () => {
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
      expect(context.document.getSelection).toHaveBeenCalledOnce();
      expect(selection.load).toHaveBeenCalledWith("text");
      expect(selection.paragraphs.load).toHaveBeenCalledWith(
        "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent",
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

      expect(selection.load).toHaveBeenCalledWith("text");
      expect(selection.paragraphs.load).toHaveBeenCalledWith(
        "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent",
      );
      expect(body.paragraphs.load).toHaveBeenCalledWith(
        "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent",
      );
      expect(body.load).not.toHaveBeenCalled();
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
            makeParagraph("Segundo párrafo con sangría.", {
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
        text: "Título\n\nPárrafo actual\n\n\tSegundo párrafo con sangría.",
        isSelection: false,
      });

      expect(selection.load).toHaveBeenCalledWith("text");
      expect(selection.paragraphs.load).toHaveBeenCalledWith(
        "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent",
      );
      expect(body.paragraphs.load).toHaveBeenCalledWith(
        "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent",
      );
      expect(body.load).not.toHaveBeenCalled();
      expect(context.sync).toHaveBeenCalledTimes(2);
    });

    it("returns an empty body text unchanged when there is no selection and the document is empty", async () => {
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

    it("keeps a Word title paragraph separated from the first body paragraph", async () => {
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

    it("propagates Word.run errors", async () => {
      const error = new Error("Office host unavailable");
      const run = installRejectingWord(error);

      await expect(adapter.getTextToAnalyze()).rejects.toThrow(
        "Office host unavailable",
      );

      expect(run).toHaveBeenCalledOnce();
    });
  });

  describe("getAppliedOriginalTexts", () => {
    // SR-DG-02 — no stylistic CCs → empty Set
    it("SR-DG-02: returns an empty set when there are no stylistic: CCs", async () => {
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
      // Non-stylistic CCs must NOT have getRange called
      expect(context.document.contentControls.items[0].getRange).not.toHaveBeenCalled();
      expect(context.document.contentControls.items[1].getRange).not.toHaveBeenCalled();
      expect(context.sync).toHaveBeenCalledTimes(1);
    });

    // SR-DG-01 — both stylistic:track-change: and stylistic:comment-only: CCs
    it("SR-DG-01: collects texts from both track-change and comment-only stylistic CCs", async () => {
      const rangeTC = { load: vi.fn(), text: "originalText1" };
      const rangeCO = { load: vi.fn(), text: "originalText2" };
      const rangeOther = { load: vi.fn(), text: "should not appear" };

      const trackChangeCC = {
        tag: "stylistic:track-change:s1",
        getRange: vi.fn(() => rangeTC),
      };
      const commentOnlyCC = {
        tag: "stylistic:comment-only:s2",
        getRange: vi.fn(() => rangeCO),
      };
      const nonStylisticCC = {
        tag: "other-cc",
        getRange: vi.fn(() => rangeOther),
      };

      const context = {
        document: {
          contentControls: {
            items: [trackChangeCC, commentOnlyCC, nonStylisticCC],
            load: vi.fn(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      installWordWithContext(context);

      const result = await adapter.getAppliedOriginalTexts();

      expect(result).toEqual(new Set(["originalText1", "originalText2"]));
      expect(trackChangeCC.getRange).toHaveBeenCalledOnce();
      expect(commentOnlyCC.getRange).toHaveBeenCalledOnce();
      expect(nonStylisticCC.getRange).not.toHaveBeenCalled();
      expect(rangeTC.load).toHaveBeenCalledWith("text");
      expect(rangeCO.load).toHaveBeenCalledWith("text");
      expect(context.sync).toHaveBeenCalledTimes(2);
    });

    it("deduplicates texts when multiple stylistic CCs span the same text", async () => {
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

  describe("applySuggestions", () => {
    it("returns an empty insertion result when there are no suggestions", async () => {
      const onProgress = vi.fn();

      await expect(adapter.applySuggestions([], onProgress)).resolves.toEqual({
        successCount: 0,
        failedSuggestions: [],
      });

      expect(commandMocks.constructor).not.toHaveBeenCalled();
      expect(commandMocks.execute).not.toHaveBeenCalled();
      expect(onProgress).not.toHaveBeenCalled();
    });

    it("maps command results into successCount, failedSuggestions, and progress events", async () => {
      const first = makeSuggestion({ id: "s-1", anchor: "uno", context: "Contexto uno" });
      const second = makeSuggestion({ id: "s-2", anchor: "dos", context: "Contexto dos" });
      const onProgress = vi.fn();

      // Suggestions are applied in reverse order (end-of-doc first): second → first.
      // The first execute call (for second) succeeds; the second (for first) fails.
      commandMocks.execute
        .mockResolvedValueOnce({ success: true, commandId: "s-2" })
        .mockResolvedValueOnce({
          success: false,
          commandId: "s-1",
          error: "Texto original no encontrado",
        });

      await expect(
        adapter.applySuggestions([first, second], onProgress),
      ).resolves.toEqual({
        successCount: 1,
        failedSuggestions: [first],
      });

      // Reverse order: second is constructed and executed first, then first
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(1, second);
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(2, first);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(1, second);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(2, first);
      expect(onProgress).toHaveBeenNthCalledWith(
        1,
        "applying",
        1,
        2,
        "Aplicando sugerencia 1 de 2...",
      );
      expect(onProgress).toHaveBeenNthCalledWith(
        2,
        "applying",
        2,
        2,
        "Aplicando sugerencia 2 de 2...",
      );
      expect(warnSpy).toHaveBeenCalledOnce();
    });

    it("converts command rejections into failed suggestions and continues processing", async () => {
      const first = makeSuggestion({ id: "s-1", anchor: "uno", context: "Contexto uno" });
      const second = makeSuggestion({ id: "s-2", anchor: "dos", context: "Contexto dos" });
      const third = makeSuggestion({ id: "s-3", anchor: "tres", context: "Contexto tres" });
      const onProgress = vi.fn();

      // Suggestions are applied in reverse order: third → second → first.
      // third succeeds, second rejects, first succeeds.
      commandMocks.execute
        .mockResolvedValueOnce({ success: true, commandId: "s-3" })
        .mockRejectedValueOnce(new Error("insert failed"))
        .mockResolvedValueOnce({ success: true, commandId: "s-1" });

      await expect(
        adapter.applySuggestions([first, second, third], onProgress),
      ).resolves.toEqual({
        successCount: 2,
        failedSuggestions: [second],
      });

      expect(commandMocks.constructor).toHaveBeenCalledTimes(3);
      // Reverse order: third, second, first
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(1, third);
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(2, second);
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(3, first);
      expect(commandMocks.execute).toHaveBeenCalledTimes(3);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(1, third);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(2, second);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(3, first);
      expect(onProgress).toHaveBeenNthCalledWith(
        1,
        "applying",
        1,
        3,
        "Aplicando sugerencia 1 de 3...",
      );
      expect(onProgress).toHaveBeenNthCalledWith(
        2,
        "applying",
        2,
        3,
        "Aplicando sugerencia 2 de 3...",
      );
      expect(onProgress).toHaveBeenNthCalledWith(
        3,
        "applying",
        3,
        3,
        "Aplicando sugerencia 3 de 3...",
      );
      expect(warnSpy).toHaveBeenCalledOnce();
    });

    it("Test A — applies suggestions in reverse array order to avoid CC interference", async () => {
      // Three suggestions: A (early in doc / index 0), B (late / index 1), C (middle / index 2)
      // Backend returns them as [A, B, C]. After reverse they should be applied: C, B, A.
      const suggA = makeSuggestion({
        id: "s-a",
        anchor: "texto al inicio del doc",
        context: "Contexto texto al inicio del doc",
      });
      const suggB = makeSuggestion({
        id: "s-b",
        anchor: "texto al final del doc",
        context: "Contexto texto al final del doc",
      });
      const suggC = makeSuggestion({
        id: "s-c",
        anchor: "texto en el medio del doc",
        context: "Contexto texto en el medio del doc",
      });

      commandMocks.execute
        .mockResolvedValueOnce({ success: true, commandId: "s-c" })
        .mockResolvedValueOnce({ success: true, commandId: "s-b" })
        .mockResolvedValueOnce({ success: true, commandId: "s-a" });

      await adapter.applySuggestions([suggA, suggB, suggC]);

      // Reverse of [A, B, C] is [C, B, A] — C must be constructed and executed first
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(1, suggC);
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(2, suggB);
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(3, suggA);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(1, suggC);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(2, suggB);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(3, suggA);
    });

    it("Test B — same-paragraph interference: both suggestions succeed when applied in reverse", async () => {
      // Simulates two suggestions in the same paragraph.
      // Suggestion 1 (early in paragraph) is applied AFTER suggestion 2 (later in paragraph)
      // because we reverse. Both must succeed — no CC overlap error.
      const earlyInParagraph = makeSuggestion({
        id: "s-early",
        anchor: "apenas eran un par de pasos alrededor de ella",
        context: "Contexto apenas eran un par de pasos alrededor de ella",
      });
      const lateInParagraph = makeSuggestion({
        id: "s-late",
        anchor: "asumió que eso sucedía porque perdía el control",
        context: "Contexto asumió que eso sucedía porque perdía el control",
      });

      // Both commands succeed — the key is execution order: lateInParagraph first, then earlyInParagraph
      commandMocks.execute
        .mockResolvedValueOnce({ success: true, commandId: "s-late" })
        .mockResolvedValueOnce({ success: true, commandId: "s-early" });

      const result = await adapter.applySuggestions([
        earlyInParagraph,
        lateInParagraph,
      ]);

      expect(result).toEqual({ successCount: 2, failedSuggestions: [] });
      // lateInParagraph must run first (it appears later in the doc, so reverse order puts it first)
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(
        1,
        lateInParagraph,
      );
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(
        2,
        earlyInParagraph,
      );
    });
  });

  describe("cleanupResolvedComments", () => {
    it("delegates to CommentCleanup and returns its counts", async () => {
      cleanupMocks.cleanupResolvedComments.mockResolvedValueOnce({
        deleted: 3,
        kept: 1,
      });

      await expect(adapter.cleanupResolvedComments()).resolves.toEqual({
        deleted: 3,
        kept: 1,
      });

      expect(cleanupMocks.cleanupResolvedComments).toHaveBeenCalledOnce();
    });

    it("propagates cleanup errors", async () => {
      cleanupMocks.cleanupResolvedComments.mockRejectedValueOnce(
        new Error("cleanup failed"),
      );

      await expect(adapter.cleanupResolvedComments()).rejects.toThrow(
        "cleanup failed",
      );
    });
  });

});
