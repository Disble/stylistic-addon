import type { Suggestion } from "../../domain/types";

export type ParentCC = {
  tag: string;
  isNullObject: boolean;
  load: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
};

export type MockRange = {
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

export type ApplyCommandTestContext = {
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

/**
 * Builds a deterministic Office.js range collection mock.
 */
function createRangeCollection(items: MockRange[]): RangeCollection {
  return { items, load: vi.fn() };
}

/**
 * Creates a search mock whose responses follow a fixed call sequence.
 */
export function createSearchMock(
  sequence: MockRange[][],
): ReturnType<typeof vi.fn> {
  let index = 0;
  return vi.fn(() => createRangeCollection(sequence[index++] ?? []));
}

/**
 * Creates a reusable range fixture with configurable nested search behavior.
 */
export function createRange(options: {
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

/**
 * Builds a canonical suggestion fixture for command tests.
 */
export function makeSuggestion(
  overrides: Partial<Suggestion> = {},
): Suggestion {
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

/**
 * Installs a strict `Word.run` harness for `ApplySuggestionCommand` tests.
 */
export function installWordContext(options: {
  documentText?: string;
  contextText?: string;
  anchorText?: string;
  contextSearchSequence?: MockRange[][];
  anchorSearchSequence?: MockRange[][];
  initialTrackingMode?: string;
  insertError?: Error;
  contextRangeParentCC?: Partial<ParentCC>;
  anchorRangeParentCC?: Partial<ParentCC>;
  onSync?: (count: number) => void | Promise<void>;
} = {}): ApplyCommandTestContext {
  const anchorText = options.anchorText ?? "texto original";
  const contextText = options.contextText ?? `Contexto con ${anchorText}.`;
  const documentText = options.documentText ?? contextText;

  const cc = { tag: "", appearance: "", cannotDelete: true };
  const insertedRange = {
    insertContentControl: vi.fn(() => cc),
    insertComment: vi.fn(),
  };

  const anchorRange = createRange({
    text: anchorText,
    parentCC: options.anchorRangeParentCC,
    insertTextImpl: () => {
      if (options.insertError) {
        throw options.insertError;
      }
      return insertedRange;
    },
  });

  const bodyRange = createRange({
    text: contextText,
    searchSequence: options.anchorSearchSequence ?? [[anchorRange]],
    parentCC: options.contextRangeParentCC,
  });

  const body = {
    ...createRange({
      text: documentText,
      searchSequence: options.contextSearchSequence ?? [[bodyRange]],
    }),
    search: createSearchMock(options.contextSearchSequence ?? [[bodyRange]]),
    text: documentText,
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
