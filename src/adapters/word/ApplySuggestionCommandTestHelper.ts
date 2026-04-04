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
  paragraphs: {
    getFirst: ReturnType<typeof vi.fn>;
  };
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
    title: string;
    appearance: string;
    cannotDelete: boolean;
  };
};

function createRangeCollection(items: MockRange[]): RangeCollection {
  return { items, load: vi.fn() };
}

export function createSearchMock(
  sequence: MockRange[][],
): ReturnType<typeof vi.fn> {
  let index = 0;
  return vi.fn(() => createRangeCollection(sequence[index++] ?? []));
}

export function createRange(options: {
  text: string;
  searchSequence?: MockRange[][];
  insertTextImpl?: () => unknown;
  parentCC?: Partial<ParentCC>;
  paragraphText?: string;
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

  const paragraphRange = {
    text: options.paragraphText ?? options.text,
    load: vi.fn(),
    search: createSearchMock(options.searchSequence ?? [[]]),
  };

  const paragraph = {
    getRange: vi.fn(() => paragraphRange),
  };

  return {
    text: options.text,
    load: vi.fn(),
    search: createSearchMock(options.searchSequence ?? [[]]),
    insertText: vi.fn(options.insertTextImpl ?? (() => defaultInsertedRange)),
    insertComment: vi.fn(),
    insertContentControl: vi.fn(() => defaultCC),
    parentContentControlOrNullObject: parentCC,
    paragraphs: {
      getFirst: vi.fn(() => paragraph),
    },
  };
}

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

export function installWordContext(options: {
  documentText?: string;
  contextText?: string;
  anchorText?: string;
  contextSearchSequence?: MockRange[][];
  anchorSearchSequence?: MockRange[][];
  paragraphSearchSequence?: MockRange[][];
  initialTrackingMode?: string;
  insertError?: Error;
  contextRangeParentCC?: Partial<ParentCC>;
  anchorRangeParentCC?: Partial<ParentCC>;
  onSync?: (count: number) => void | Promise<void>;
  setupParagraphSearch?: (
    ctx: ApplyCommandTestContext,
    contextRange: MockRange,
    paragraphRangeRef: { current: MockRange | null },
    anchorRangeRef: { current: MockRange | null },
  ) => void;
  anchorRangeRef?: { current: MockRange | null };
} = {}): ApplyCommandTestContext {
  const anchorText = options.anchorText ?? "texto original";
  const contextText = options.contextText ?? `Contexto con ${anchorText}.`;
  const documentText = options.documentText ?? contextText;

  const cc = { tag: "", title: "", appearance: "", cannotDelete: true };
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

  const anchorSearchSequence = options.anchorSearchSequence ?? [[anchorRange]];
  const bodyRange = createRange({
    text: contextText,
    searchSequence: anchorSearchSequence,
    parentCC: options.contextRangeParentCC,
  });

  const paragraphSearchSeq =
    options.paragraphSearchSequence ?? anchorSearchSequence;
  const paragraphSearchMock = createSearchMock(paragraphSearchSeq);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (bodyRange.paragraphs as any).getFirst().getRange("Whole").search =
    paragraphSearchMock;

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

  anchorRange.insertContentControl = vi.fn(() => cc);
  bodyRange.insertContentControl = vi.fn(() => cc);

  const testContext: ApplyCommandTestContext = {
    context,
    bodyRange,
    anchorRange,
    insertedRange,
    cc,
  };

  options.setupParagraphSearch?.(
    testContext,
    bodyRange,
    { current: null },
    options.anchorRangeRef ?? { current: null },
  );

  return testContext;
}
