import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { Suggestion } from "../../domain/types";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";

type ParentCC = {
  tag: string;
  isNullObject: boolean;
  load: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
};

type MockRange = {
  text: string;
  load: ReturnType<typeof vi.fn>;
  search: ReturnType<typeof vi.fn>;
  insertText: ReturnType<typeof vi.fn>;
  insertComment: ReturnType<typeof vi.fn>;
  insertContentControl: ReturnType<typeof vi.fn>;
  parentContentControlOrNullObject: ParentCC;
};

type RangeCollection = {
  items: MockRange[];
  load: ReturnType<typeof vi.fn>;
};

type ApplyTestContext = {
  context: {
    document: {
      body: MockRange & { search: ReturnType<typeof vi.fn>; text: string };
      load: ReturnType<typeof vi.fn>;
      changeTrackingMode: string;
    };
    sync: ReturnType<typeof vi.fn>;
  };
  bodyRange: MockRange;
  anchorRange: MockRange;
  insertedRange: {
    insertContentControl: ReturnType<typeof vi.fn>;
    insertComment: ReturnType<typeof vi.fn>;
  };
  cc: {
    tag: string;
    appearance: string;
    cannotDelete: boolean;
  };
};

function createRangeCollection(items: MockRange[]): RangeCollection {
  return { items, load: vi.fn() };
}

function createSearchMock(sequence: MockRange[][]): ReturnType<typeof vi.fn> {
  let index = 0;
  return vi.fn(() => createRangeCollection(sequence[index++] ?? []));
}

function createRange(options: {
  text: string;
  searchSequence?: MockRange[][];
  insertTextImpl?: () => unknown;
  parentCC?: Partial<ParentCC>;
}): MockRange {
  const parentCC: ParentCC = {
    tag: "",
    isNullObject: true,
    load: vi.fn(),
    delete: vi.fn(),
    ...options.parentCC,
  };

  const defaultCC = { tag: "", appearance: "", cannotDelete: true };
  const defaultInsertedRange = {
    insertContentControl: vi.fn(() => defaultCC),
    insertComment: vi.fn(),
  };

  return {
    text: options.text,
    load: vi.fn(),
    search: createSearchMock(options.searchSequence ?? [[]]),
    insertText: vi.fn(options.insertTextImpl ?? (() => defaultInsertedRange)),
    insertComment: vi.fn(),
    insertContentControl: vi.fn(() => defaultCC),
    parentContentControlOrNullObject: parentCC,
  };
}

function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  const anchor = overrides.anchor ?? "texto original";
  return {
    id: "s1",
    context: overrides.context ?? `Contexto con ${anchor}.`,
    anchor,
    suggestedText: "texto sugerido",
    justification: "Mejora la claridad",
    category: "Estilo",
    severity: "medium",
    type: "track-change",
    ...overrides,
  };
}

function installWordContext(options: {
  bodyText?: string;
  contextSearchSequence?: MockRange[][];
  anchorSearchSequence?: MockRange[][];
  initialTrackingMode?: string;
  insertError?: Error;
  contextRangeParentCC?: Partial<ParentCC>;
  anchorRangeParentCC?: Partial<ParentCC>;
  onSync?: (count: number) => void | Promise<void>;
} = {}): ApplyTestContext {
  const cc = { tag: "", appearance: "", cannotDelete: true };
  const insertedRange = {
    insertContentControl: vi.fn(() => cc),
    insertComment: vi.fn(),
  };

  const anchorRange = createRange({
    text: options.bodyText ?? "texto original",
    parentCC: options.anchorRangeParentCC,
    insertTextImpl: () => {
      if (options.insertError) {
        throw options.insertError;
      }
      return insertedRange;
    },
  });

  const bodyRange = createRange({
    text: options.bodyText ?? `Contexto con texto original.`,
    searchSequence: options.anchorSearchSequence ?? [[anchorRange]],
    parentCC: options.contextRangeParentCC,
  });

  const body = {
    ...createRange({
      text: options.bodyText ?? `Contexto con texto original.`,
      searchSequence: options.contextSearchSequence ?? [[bodyRange]],
    }),
    search: createSearchMock(options.contextSearchSequence ?? [[bodyRange]]),
    text: options.bodyText ?? `Contexto con texto original.`,
  };

  let syncCount = 0;
  const context = {
    document: {
      body,
      load: vi.fn(),
      changeTrackingMode: options.initialTrackingMode ?? "off",
    },
    sync: vi.fn(async () => {
      syncCount += 1;
      await options.onSync?.(syncCount);
    }),
  };

  vi.stubGlobal("Word", {
    ChangeTrackingMode: {
      off: "off",
      trackAll: "trackAll",
      trackMine: "trackMine",
    },
    InsertLocation: {
      replace: "Replace",
    },
    run: vi.fn(async (callback: (ctx: typeof context) => unknown) =>
      callback(context),
    ),
  });

  return { context, bodyRange, anchorRange, insertedRange, cc };
}

describe("ApplySuggestionCommand", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
    vi.spyOn(console, "warn").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("returns a distinct error when context is not found", async () => {
    installWordContext({ contextSearchSequence: [[], [], []] });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor no encontrado en el contexto",
    });
  });

  it("returns a distinct error when anchor is not found inside the located context", async () => {
    installWordContext({
      contextSearchSequence: [[createRange({ text: "Contexto con texto original." })]],
      anchorSearchSequence: [[], [], []],
    });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Anchor no encontrado en el contexto",
    });
  });

  it("uses two-step search: body finds context, then context range finds anchor", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.context.document.body.search).toHaveBeenCalledWith(
      "Contexto con texto original.",
      { matchCase: true, matchWholeWord: false },
    );
    expect(env.bodyRange.search).toHaveBeenCalledWith("texto original", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(env.anchorRange.insertText).toHaveBeenCalledWith(
      "texto sugerido",
      "Replace",
    );
  });

  it("tries ignorePunct+ignoreSpace when exact anchor search fails", async () => {
    const context = installWordContext();
    const bodyRange = context.bodyRange;
    context.bodyRange.search = createSearchMock([
      [],
      [context.anchorRange],
    ]);

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(context.bodyRange.search).toHaveBeenNthCalledWith(1, "texto original", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(context.bodyRange.search).toHaveBeenNthCalledWith(2, "texto original", {
      matchCase: true,
      matchWholeWord: false,
      ignorePunct: true,
      ignoreSpace: true,
    });
  });

  it("supports comment-only suggestions through the shared anchor resolver", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion({ type: "comment-only", suggestedText: undefined }),
    ).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.context.document.changeTrackingMode).toBe("off");
    expect(env.anchorRange.insertComment).toHaveBeenCalledWith(
      "[Estilo]\nMejora la claridad",
    );
  });

  it("re-resolves the anchor after removing an existing content control", async () => {
    const coveredParentCC: ParentCC = {
      tag: "stylistic:track-change:s1",
      isNullObject: false,
      load: vi.fn(),
      delete: vi.fn(),
    };
    const coveredAnchor = createRange({
      text: "texto original",
      parentCC: coveredParentCC,
    });
    const coveredContext = createRange({
      text: "Contexto con texto original.",
      searchSequence: [[coveredAnchor]],
    });
    const freshAnchor = createRange({ text: "texto original" });
    const freshContext = createRange({
      text: "Contexto con texto original.",
      searchSequence: [[freshAnchor]],
    });

    const env = installWordContext({
      contextSearchSequence: [[coveredContext], [freshContext]],
    });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(coveredParentCC.delete).toHaveBeenCalledWith(true);
    expect(env.context.document.body.search).toHaveBeenCalledTimes(2);
    expect(freshAnchor.insertText).toHaveBeenCalled();
  });

  it("rejects insert-only suggestions without touching Word", async () => {
    const env = installWordContext();

    const result = await new ApplySuggestionCommand(
      makeSuggestion({ anchor: "", context: "", suggestedText: "texto sugerido" }),
    ).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Insert-only suggestions require anchor text",
    });
    expect(env.context.document.body.search).not.toHaveBeenCalled();
  });

  it("restores changeTrackingMode when insertText throws", async () => {
    const env = installWordContext({
      initialTrackingMode: "off",
      insertError: new Error("insert failed"),
    });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "insert failed",
    });
    expect(env.context.document.changeTrackingMode).toBe("off");
    expect(env.context.document.load).toHaveBeenCalledWith("changeTrackingMode");
  });

  it("loads changeTrackingMode before reading the previous mode", async () => {
    const env = installWordContext({ initialTrackingMode: "trackMine" });

    const result = await new ApplySuggestionCommand(makeSuggestion()).execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(env.context.document.load).toHaveBeenCalledWith("changeTrackingMode");
    expect(env.context.document.changeTrackingMode).toBe("trackMine");
  });
});
