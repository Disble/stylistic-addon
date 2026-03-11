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
      const selection = { load: vi.fn(), text: "Texto seleccionado" };
      const body = { load: vi.fn(), text: "Texto del documento" };
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
      expect(body.load).not.toHaveBeenCalled();
      expect(context.sync).toHaveBeenCalledTimes(1);
    });

    it("falls back to the full document body when selection is empty or whitespace", async () => {
      const selection = { load: vi.fn(), text: "   \n  " };
      const body = { load: vi.fn(), text: "Texto completo del documento" };
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
      expect(body.load).toHaveBeenCalledWith("text");
      expect(context.sync).toHaveBeenCalledTimes(2);
    });

    it("returns an empty body text unchanged when there is no selection and the document is empty", async () => {
      const selection = { load: vi.fn(), text: "" };
      const body = { load: vi.fn(), text: "" };
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
});
