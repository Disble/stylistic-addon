import type { Suggestion, CommandResult } from "../../domain/types";

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
  OVERLAPPING_RELATIONS: ["Equal", "Contains", "ContainsStart", "ContainsEnd", "Inside", "InsideStart", "InsideEnd", "OverlapsBefore", "OverlapsAfter"],
  COMMENT_ONLY_TAG_PREFIX: "stylistic:comment-only:",
}));

import { WordAdapter } from "./WordAdapter";

type WordRunCallback<T> = (context: any) => Promise<T> | T;

function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  return {
    id: "s-1",
    originalText: "texto original",
    suggestedText: "texto sugerido",
    justification: "Mas claro",
    category: "Claridad",
    severity: "medium",
    ...overrides,
  };
}

function installWordWithContext(context: any) {
  const run = vi.fn(async <T>(callback: WordRunCallback<T>) => callback(context));
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
  }> = {}
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
    cleanupMocks.cleanupResolvedComments.mockResolvedValue({ deleted: 0, kept: 0 });

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

      expect(selection.load).toHaveBeenCalledWith("text");
      expect(selection.paragraphs.load).toHaveBeenCalledWith(
        "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent"
      );
      expect(body.paragraphs.load).toHaveBeenCalledWith(
        "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent"
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

      expect(selection.load).toHaveBeenCalledWith("text");
      expect(selection.paragraphs.load).toHaveBeenCalledWith(
        "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent"
      );
      expect(body.paragraphs.load).toHaveBeenCalledWith(
        "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent"
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

      await expect(adapter.getTextToAnalyze()).resolves.toEqual({ text: "", isSelection: false });
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
            makeParagraph("Segundo párrafo con sangría visual.", { firstLineIndent: 18 }),
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

      await expect(adapter.getTextToAnalyze()).rejects.toThrow("Office host unavailable");

      expect(run).toHaveBeenCalledOnce();
    });
  });

  describe("getAppliedOriginalTexts", () => {
    it("returns an empty set when there are no Stylistic deleted tracked changes", async () => {
      const tracked = {
        items: [
          { author: "Stylistic", type: "Inserted", getRange: vi.fn() },
          { author: "Otro", type: "Deleted", getRange: vi.fn() },
        ],
        load: vi.fn(),
      };
      const context = {
        document: {
          body: {
            getTrackedChanges: vi.fn(() => tracked),
          },
          contentControls: { items: [], load: vi.fn() },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      installWordWithContext(context);

      const result = await adapter.getAppliedOriginalTexts();

      expect(result).toEqual(new Set());
      expect(tracked.load).toHaveBeenCalledWith({ select: "author,type" });
      expect(tracked.items[0].getRange).not.toHaveBeenCalled();
      expect(tracked.items[1].getRange).not.toHaveBeenCalled();
      expect(context.sync).toHaveBeenCalledTimes(1);
    });

    it("collects and deduplicates texts from Stylistic deleted tracked changes only", async () => {
      const rangeA = { load: vi.fn(), text: "primero" };
      const rangeB = { load: vi.fn(), text: "segundo" };
      const rangeDuplicate = { load: vi.fn(), text: "primero" };
      const ignoredInsertedRange = { load: vi.fn(), text: "ignorado" };
      const ignoredAuthorRange = { load: vi.fn(), text: "ignorado" };

      const deletedA = { author: "Stylistic", type: "Deleted", getRange: vi.fn(() => rangeA) };
      const deletedB = { author: "Stylistic", type: "Deleted", getRange: vi.fn(() => rangeB) };
      const deletedDuplicate = {
        author: "Stylistic",
        type: "Deleted",
        getRange: vi.fn(() => rangeDuplicate),
      };
      const inserted = {
        author: "Stylistic",
        type: "Inserted",
        getRange: vi.fn(() => ignoredInsertedRange),
      };
      const otherAuthor = {
        author: "Reviewer",
        type: "Deleted",
        getRange: vi.fn(() => ignoredAuthorRange),
      };
      const tracked = {
        items: [deletedA, inserted, otherAuthor, deletedB, deletedDuplicate],
        load: vi.fn(),
      };
      const context = {
        document: {
          body: {
            getTrackedChanges: vi.fn(() => tracked),
          },
          contentControls: { items: [], load: vi.fn() },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      installWordWithContext(context);

      const result = await adapter.getAppliedOriginalTexts();

      expect(result).toEqual(new Set(["primero", "segundo"]));
      expect(deletedA.getRange).toHaveBeenCalledOnce();
      expect(deletedB.getRange).toHaveBeenCalledOnce();
      expect(deletedDuplicate.getRange).toHaveBeenCalledOnce();
      expect(inserted.getRange).not.toHaveBeenCalled();
      expect(otherAuthor.getRange).not.toHaveBeenCalled();
      expect(rangeA.load).toHaveBeenCalledWith("text");
      expect(rangeB.load).toHaveBeenCalledWith("text");
      expect(rangeDuplicate.load).toHaveBeenCalledWith("text");
      expect(context.sync).toHaveBeenCalledTimes(2);
    });

    it("includes the range text of active comment-only CCs", async () => {
      const ccRange = { load: vi.fn(), text: "texto en observación" };
      const commentOnlyCC = {
        tag: "stylistic:comment-only:chunk0-0",
        getRange: vi.fn(() => ccRange),
      };

      const tracked = { items: [], load: vi.fn() };
      const context = {
        document: {
          body: {
            getTrackedChanges: vi.fn(() => tracked),
          },
          contentControls: {
            items: [commentOnlyCC],
            load: vi.fn(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      installWordWithContext(context);

      const result = await adapter.getAppliedOriginalTexts();

      expect(result).toEqual(new Set(["texto en observación"]));
      expect(commentOnlyCC.getRange).toHaveBeenCalledOnce();
      expect(ccRange.load).toHaveBeenCalledWith("text");
      expect(context.sync).toHaveBeenCalledTimes(2);
    });

    it("propagates Word.run errors", async () => {
      installRejectingWord(new Error("Tracked changes unavailable"));

      await expect(adapter.getAppliedOriginalTexts()).rejects.toThrow(
        "Tracked changes unavailable"
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
      const first = makeSuggestion({ id: "s-1", originalText: "uno" });
      const second = makeSuggestion({ id: "s-2", originalText: "dos" });
      const onProgress = vi.fn();

      commandMocks.execute
        .mockResolvedValueOnce({ success: true, commandId: "s-1" })
        .mockResolvedValueOnce({
          success: false,
          commandId: "s-2",
          error: "Texto original no encontrado",
        });

      await expect(adapter.applySuggestions([first, second], onProgress)).resolves.toEqual({
        successCount: 1,
        failedSuggestions: [second],
      });

      expect(commandMocks.constructor).toHaveBeenNthCalledWith(1, first);
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(2, second);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(1, first);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(2, second);
      expect(onProgress).toHaveBeenNthCalledWith(
        1,
        "applying",
        1,
        2,
        "Aplicando sugerencia 1 de 2..."
      );
      expect(onProgress).toHaveBeenNthCalledWith(
        2,
        "applying",
        2,
        2,
        "Aplicando sugerencia 2 de 2..."
      );
      expect(warnSpy).toHaveBeenCalledOnce();
    });

    it("converts command rejections into failed suggestions and continues processing", async () => {
      const first = makeSuggestion({ id: "s-1", originalText: "uno" });
      const second = makeSuggestion({ id: "s-2", originalText: "dos" });
      const third = makeSuggestion({ id: "s-3", originalText: "tres" });
      const onProgress = vi.fn();

      commandMocks.execute
        .mockResolvedValueOnce({ success: true, commandId: "s-1" })
        .mockRejectedValueOnce(new Error("insert failed"))
        .mockResolvedValueOnce({ success: true, commandId: "s-3" });

      await expect(adapter.applySuggestions([first, second, third], onProgress)).resolves.toEqual({
        successCount: 2,
        failedSuggestions: [second],
      });

      expect(commandMocks.constructor).toHaveBeenCalledTimes(3);
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(1, first);
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(2, second);
      expect(commandMocks.constructor).toHaveBeenNthCalledWith(3, third);
      expect(commandMocks.execute).toHaveBeenCalledTimes(3);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(1, first);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(2, second);
      expect(commandMocks.execute).toHaveBeenNthCalledWith(3, third);
      expect(onProgress).toHaveBeenNthCalledWith(
        1,
        "applying",
        1,
        3,
        "Aplicando sugerencia 1 de 3..."
      );
      expect(onProgress).toHaveBeenNthCalledWith(
        2,
        "applying",
        2,
        3,
        "Aplicando sugerencia 2 de 3..."
      );
      expect(onProgress).toHaveBeenNthCalledWith(
        3,
        "applying",
        3,
        3,
        "Aplicando sugerencia 3 de 3..."
      );
      expect(warnSpy).toHaveBeenCalledOnce();
    });
  });

  describe("cleanupResolvedComments", () => {
    it("delegates to CommentCleanup and returns its counts", async () => {
      cleanupMocks.cleanupResolvedComments.mockResolvedValueOnce({ deleted: 3, kept: 1 });

      await expect(adapter.cleanupResolvedComments()).resolves.toEqual({ deleted: 3, kept: 1 });

      expect(cleanupMocks.cleanupResolvedComments).toHaveBeenCalledOnce();
    });

    it("propagates cleanup errors", async () => {
      cleanupMocks.cleanupResolvedComments.mockRejectedValueOnce(new Error("cleanup failed"));

      await expect(adapter.cleanupResolvedComments()).rejects.toThrow("cleanup failed");
    });
  });

  // Helper: builds a context mock using Content Controls
  function makeResolveSuggestionContext({
    ccFound = true,
    spanTCItems = [] as any[],
    comments = [] as any[],
  }) {
    const spanTCCollection = { items: spanTCItems, load: vi.fn() };
    
    const cc = {
      getTrackedChanges: vi.fn(() => spanTCCollection),
      getRange: vi.fn(() => ({ compareLocationWith: vi.fn() })),
      delete: vi.fn(),
    };

    const ccsCollection = {
      items: ccFound ? [cc] : [],
      load: vi.fn(),
    };

    const commentsCollection = { items: comments, load: vi.fn() };

    return {
      document: {
        contentControls: {
          getByTag: vi.fn(() => ccsCollection),
        },
        body: {
          getComments: vi.fn(() => commentsCollection),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
      _ccsCollection: ccsCollection,
      _commentsCollection: commentsCollection,
      _cc: cc,
    };
  }

  describe("acceptSuggestion", () => {
    it("3.1 - happy path: accepts 2 Stylistic TCs (Deleted+Added) and deletes colocated comment", async () => {
      const suggestion = makeSuggestion({ id: "s-1", originalText: "texto original" });

      const tcAccept1 = vi.fn();
      const tcAccept2 = vi.fn();
      const commentDeleteSpy = vi.fn();

      const spanTCItems = [
        { author: "Stylistic", type: "Deleted", accept: tcAccept1, reject: vi.fn() },
        { author: "Stylistic", type: "Added", accept: tcAccept2, reject: vi.fn() },
      ];

      const commentRange = { compareLocationWith: vi.fn(() => ({ value: "Equal" })) };
      const comment = {
        authorName: "Stylistic",
        getRange: vi.fn(() => commentRange),
        delete: commentDeleteSpy,
      };

      const context = makeResolveSuggestionContext({
        ccFound: true,
        spanTCItems,
        comments: [comment],
      });

      installWordWithContext(context);

      const result = await adapter.acceptSuggestion(suggestion);

      expect(result.status).toBe("accepted");
      expect(result.trackedChangesAffected).toBe(2);
      expect(result.commentDeleted).toBe(true);
      expect(tcAccept1).toHaveBeenCalledOnce();
      expect(tcAccept2).toHaveBeenCalledOnce();
      expect(commentDeleteSpy).toHaveBeenCalledOnce();
      expect(context._cc.delete).toHaveBeenCalledWith(true);
    });

    it("3.2 - already-resolved: Content Control not found", async () => {
      const suggestion = makeSuggestion({ originalText: "texto original" });

      const context = makeResolveSuggestionContext({ ccFound: false });
      installWordWithContext(context);

      const result = await adapter.acceptSuggestion(suggestion);

      expect(result.status).toBe("already-resolved");
      expect(result.trackedChangesAffected).toBe(0);
      expect(result.commentDeleted).toBe(false);
    });

    it("3.4 - already-resolved: Content Control found but has NO Stylistic TCs", async () => {
      const suggestion = makeSuggestion({ originalText: "texto original" });

      const otherTC = {
        author: "OtherUser",
        type: "Deleted",
        accept: vi.fn(),
        reject: vi.fn(),
      } as any;

      const context = makeResolveSuggestionContext({
        ccFound: true,
        spanTCItems: [otherTC]
      });
      installWordWithContext(context);

      const result = await adapter.acceptSuggestion(suggestion);

      expect(result.status).toBe("already-resolved");
      expect(result.trackedChangesAffected).toBe(0);
      expect(context._cc.delete).toHaveBeenCalledWith(true);
    });

    it("3.5 - TCs found but no comment (comment already deleted)", async () => {
      const suggestion = makeSuggestion({ originalText: "texto original" });
      const tcAcceptSpy = vi.fn();

      const spanTCItems = [
        { author: "Stylistic", type: "Deleted", accept: tcAcceptSpy, reject: vi.fn() },
      ];

      const context = makeResolveSuggestionContext({
        ccFound: true,
        spanTCItems,
        comments: [], // no comments
      });
      installWordWithContext(context);

      const result = await adapter.acceptSuggestion(suggestion);

      expect(result.status).toBe("accepted");
      expect(result.trackedChangesAffected).toBe(1);
      expect(result.commentDeleted).toBe(false);
      expect(result.error).toBeUndefined();
      expect(tcAcceptSpy).toHaveBeenCalledOnce();
      expect(context._cc.delete).toHaveBeenCalledWith(true);
    });

    it("3.6 - Word.run throws: returns error status without throwing", async () => {
      const suggestion = makeSuggestion({ originalText: "texto original" });

      installRejectingWord(new Error("Document is read-only"));

      const result = await adapter.acceptSuggestion(suggestion);

      expect(result.status).toBe("error");
      expect(result.error).toContain("Document is read-only");
    });

    it("3.9 - comment-only accept: deletes comment and CC, returns accepted with 0 TCs", async () => {
      const suggestion = makeSuggestion({
        id: "s-co-1",
        type: "comment-only",
        suggestedText: undefined,
      });

      const commentDeleteSpy = vi.fn();
      const commentRange = { compareLocationWith: vi.fn(() => ({ value: "Equal" })) };
      const comment = {
        authorName: "Stylistic",
        getRange: vi.fn(() => commentRange),
        delete: commentDeleteSpy,
      };

      const context = makeResolveSuggestionContext({
        ccFound: true,
        spanTCItems: [],
        comments: [comment],
      });

      installWordWithContext(context);

      const result = await adapter.acceptSuggestion(suggestion);

      expect(result.status).toBe("accepted");
      expect(result.trackedChangesAffected).toBe(0);
      expect(result.commentDeleted).toBe(true);
      expect(commentDeleteSpy).toHaveBeenCalledOnce();
      expect(context._cc.delete).toHaveBeenCalledWith(true);
    });
  });

  describe("rejectSuggestion", () => {
    it("3.7 - happy path: rejects 2 Stylistic TCs and deletes colocated comment", async () => {
      const suggestion = makeSuggestion({ id: "s-1", originalText: "texto original" });

      const tcReject1 = vi.fn();
      const tcReject2 = vi.fn();
      const tcAccept1 = vi.fn();
      const tcAccept2 = vi.fn();
      const commentDeleteSpy = vi.fn();

      const spanTCItems = [
        { author: "Stylistic", type: "Deleted", accept: tcAccept1, reject: tcReject1 },
        { author: "Stylistic", type: "Added", accept: tcAccept2, reject: tcReject2 },
      ];

      const commentRange = { compareLocationWith: vi.fn(() => ({ value: "Equal" })) };
      const comment = {
        authorName: "Stylistic",
        getRange: vi.fn(() => commentRange),
        delete: commentDeleteSpy,
      };

      const context = makeResolveSuggestionContext({
        ccFound: true,
        spanTCItems,
        comments: [comment],
      });
      installWordWithContext(context);

      const result = await adapter.rejectSuggestion(suggestion);

      expect(result.status).toBe("rejected");
      expect(result.trackedChangesAffected).toBe(2);
      expect(tcReject1).toHaveBeenCalledOnce();
      expect(tcReject2).toHaveBeenCalledOnce();
      expect(tcAccept1).not.toHaveBeenCalled();
      expect(tcAccept2).not.toHaveBeenCalled();
      expect(commentDeleteSpy).toHaveBeenCalledOnce();
      expect(context._cc.delete).toHaveBeenCalledWith(true);
    });

    it("3.8 - already-resolved: Content Control not found", async () => {
      const suggestion = makeSuggestion({ originalText: "texto original" });

      const context = makeResolveSuggestionContext({ ccFound: false });
      installWordWithContext(context);

      const result = await adapter.rejectSuggestion(suggestion);

      expect(result.status).toBe("already-resolved");
    });

    it("3.10 - comment-only reject: deletes comment and CC, returns rejected with 0 TCs", async () => {
      const suggestion = makeSuggestion({
        id: "s-co-2",
        type: "comment-only",
        suggestedText: undefined,
      });

      const commentDeleteSpy = vi.fn();
      const commentRange = { compareLocationWith: vi.fn(() => ({ value: "Equal" })) };
      const comment = {
        authorName: "Stylistic",
        getRange: vi.fn(() => commentRange),
        delete: commentDeleteSpy,
      };

      const context = makeResolveSuggestionContext({
        ccFound: true,
        spanTCItems: [],
        comments: [comment],
      });

      installWordWithContext(context);

      const result = await adapter.rejectSuggestion(suggestion);

      expect(result.status).toBe("rejected");
      expect(result.trackedChangesAffected).toBe(0);
      expect(result.commentDeleted).toBe(true);
      expect(commentDeleteSpy).toHaveBeenCalledOnce();
      expect(context._cc.delete).toHaveBeenCalledWith(true);
    });
  });
});
