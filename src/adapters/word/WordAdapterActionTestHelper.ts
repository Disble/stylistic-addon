import type { Suggestion } from "../../domain/types";

type MockTrackedChange = {
  id?: string;
  type?: string;
  accept?: ReturnType<typeof vi.fn>;
  reject?: ReturnType<typeof vi.fn>;
  getRange?: ReturnType<typeof vi.fn>;
};

type MockComment = {
  authorName?: string;
  getRange: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
};

type ResolveSuggestionContext = {
  document: {
    contentControls: {
      getByTag: ReturnType<typeof vi.fn>;
    };
    body: {
      getComments: ReturnType<typeof vi.fn>;
      getTrackedChanges: ReturnType<typeof vi.fn>;
    };
  };
  sync: ReturnType<typeof vi.fn>;
  _ccsCollection: { items: unknown[]; load: ReturnType<typeof vi.fn> };
  _commentsCollection: { items: MockComment[]; load: ReturnType<typeof vi.fn> };
  _bodyTCCollection: { items: MockTrackedChange[]; load: ReturnType<typeof vi.fn> };
  _cc: {
    getTrackedChanges: ReturnType<typeof vi.fn>;
    getRange: ReturnType<typeof vi.fn>;
    delete: ReturnType<typeof vi.fn>;
  };
};

type MockWordGlobal = {
  run: <T>(callback: (ctx: ResolveSuggestionContext) => Promise<T> | T) => Promise<T>;
};

/**
 * Builds a canonical suggestion fixture for WordAdapter action tests.
 */
export function makeSuggestion(
  overrides: Partial<Suggestion> = {},
): Suggestion {
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

/**
 * Installs a `Word.run` mock that executes the callback with the provided context.
 */
export function installWordWithContext(context: ResolveSuggestionContext) {
  const run: MockWordGlobal["run"] = async <T>(
    callback: (ctx: ResolveSuggestionContext) => Promise<T> | T
  ) => callback(context);
  vi.stubGlobal("Word", { run } satisfies MockWordGlobal);
  return run;
}

/**
 * Installs a rejecting `Word.run` mock for sad-path adapter tests.
 */
export function installRejectingWord(error: Error) {
  const run: MockWordGlobal["run"] = async () => {
    throw error;
  };
  vi.stubGlobal("Word", { run } satisfies MockWordGlobal);
  return run;
}

/**
 * Builds a resolve-suggestion context centered on a Content Control anchor.
 *
 * This helper intentionally supports asymmetry between `cc.getTrackedChanges()`
 * and `body.getTrackedChanges()` because real Word behavior may expose only one
 * side of a replace operation through the CC-scoped collection.
 */
export function makeResolveSuggestionContext({
  ccFound = true,
  spanTCItems = [] as MockTrackedChange[],
  bodyTCItems = [] as MockTrackedChange[],
  bodyTCRelations = [] as string[],
  comments = [] as MockComment[],
}): ResolveSuggestionContext {
  const spanTCCollection = { items: spanTCItems, load: vi.fn() };
  const bodyTCCollection = { items: bodyTCItems, load: vi.fn() };

  const cc = {
    getTrackedChanges: vi.fn(() => spanTCCollection),
    getRange: vi.fn(() => ({ compareLocationWith: vi.fn() })),
    delete: vi.fn(),
  };

  const bodyTrackedChangeRanges = bodyTCItems.map((_, index) => ({
    compareLocationWith: vi.fn(() => ({
      value: bodyTCRelations[index] ?? "Disjoint",
    })),
  }));

  for (let index = 0; index < bodyTCItems.length; index += 1) {
    const trackedChange = bodyTCItems[index];
    trackedChange.getRange = vi.fn(() => bodyTrackedChangeRanges[index]);
  }

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
        getTrackedChanges: vi.fn(() => bodyTCCollection),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
    _ccsCollection: ccsCollection,
    _commentsCollection: commentsCollection,
    _bodyTCCollection: bodyTCCollection,
    _cc: cc,
  };
}
