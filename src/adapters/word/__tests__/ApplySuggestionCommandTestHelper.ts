import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import type {
  ApplyCommandTestContext,
  MockRange,
  ParentCC,
  RangeCollection,
} from "./ApplySuggestionCommandTestHelper.types";

function createRangeCollection(items: MockRange[]): RangeCollection {
  return { items, load: vi.fn() };
}

function createClientResult<T>(value: T) {
  return { value };
}

/** Creates a deterministic Word search mock that returns one range batch per call. */
export function createSearchMock(sequence: MockRange[][]): ReturnType<typeof vi.fn> {
  let index = 0;
  return vi.fn(() => createRangeCollection(sequence[index++] ?? []));
}

/** Creates a fake Word range with configurable reviewed text and search behavior. */
export function createRange(options: {
  text: string;
  searchSequence?: MockRange[][];
  insertTextImpl?: (text: string, insertLocation: string) => unknown;
  parentCC?: Partial<ParentCC>;
  paragraphText?: string;
  reviewedCurrentText?: string;
  reviewedOriginalText?: string;
}): MockRange {
  const parentCC: ParentCC = {
    tag: "",
    isNullObject: true,
    load: vi.fn(),
    delete: vi.fn(),
    ...options.parentCC,
  };

  const defaultCC = { tag: "", appearance: "", cannotDelete: true };
  const defaultInsertedParagraphRange = {
    text: options.paragraphText ?? options.text,
    load: vi.fn(),
    search: createSearchMock([[]]),
  };
  const defaultInsertedRange = {
    text: "texto sugerido",
    load: vi.fn(),
    search: createSearchMock([[]]),
    getReviewedText: vi.fn((version: "Current" | "Original") =>
      createClientResult(version === "Original" ? "" : "texto sugerido")
    ),
    paragraphs: {
      getFirst: vi.fn(() => ({
        getRange: vi.fn(() => defaultInsertedParagraphRange),
      })),
    },
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
    font: { italic: false, bold: false },
    load: vi.fn(),
    search: createSearchMock(options.searchSequence ?? [[]]),
    insertText: vi.fn(
      options.insertTextImpl ??
        ((text: string) => {
          defaultInsertedRange.text = text;
          defaultInsertedParagraphRange.text = options.paragraphText ?? options.text;
          defaultInsertedRange.getReviewedText = vi.fn((version: "Current" | "Original") =>
            createClientResult(version === "Original" ? "" : text)
          );
          return defaultInsertedRange;
        })
    ),
    getReviewedText: vi.fn((version: "Current" | "Original") =>
      createClientResult(
        version === "Original"
          ? (options.reviewedOriginalText ?? "")
          : (options.reviewedCurrentText ?? options.text)
      )
    ),
    insertComment: vi.fn(),
    insertContentControl: vi.fn(() => defaultCC),
    parentContentControlOrNullObject: parentCC,
    paragraphs: {
      getFirst: vi.fn(() => paragraph),
    },
  };
}

/** Builds a default suggestion fixture for apply-suggestion command tests. */
export function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
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

/** Installs a fake Word context tailored for apply-suggestion command scenarios. */
export function installWordContext(
  options: {
    documentText?: string;
    contextText?: string;
    anchorText?: string;
    contextSearchSequence?: MockRange[][];
    anchorSearchSequence?: MockRange[][];
    paragraphSearchSequence?: MockRange[][];
    initialTrackingMode?: string;
    insertError?: Error;
    insertedRange?: {
      text?: string;
      reviewedCurrentText?: string;
      reviewedOriginalText?: string;
      searchSequence?: MockRange[][];
      paragraphText?: string;
    };
    contextRangeParentCC?: Partial<ParentCC>;
    anchorRangeParentCC?: Partial<ParentCC>;
    onSync?: (count: number) => void | Promise<void>;
    setupParagraphSearch?: (
      ctx: ApplyCommandTestContext,
      contextRange: MockRange,
      paragraphRangeRef: { current: MockRange | null },
      anchorRangeRef: { current: MockRange | null }
    ) => void;
    anchorRangeRef?: { current: MockRange | null };
    useOperationalWrapper?: boolean;
  } = {}
): ApplyCommandTestContext {
  const anchorText = options.anchorText ?? "texto original";
  const contextText = options.contextText ?? `Contexto con ${anchorText}.`;
  const documentText = options.documentText ?? contextText;

  const cc = { tag: "", title: "", appearance: "", cannotDelete: true };
  const operationalWrapperRange = createRange({
    text: contextText,
    paragraphText: contextText,
  });
  const operationalWrapper = {
    tag: "",
    title: "",
    appearance: "",
    cannotDelete: true,
    getRange: vi.fn(() => operationalWrapperRange),
  };
  let insertedCurrentText = options.insertedRange?.reviewedCurrentText;
  const insertedRange = createRange({
    text: options.insertedRange?.text ?? options.documentText ?? contextText,
    reviewedCurrentText: insertedCurrentText ?? "texto sugerido",
    reviewedOriginalText: options.insertedRange?.reviewedOriginalText ?? "",
    searchSequence: options.insertedRange?.searchSequence ?? [[]],
    paragraphText:
      options.insertedRange?.paragraphText ??
      options.insertedRange?.text ??
      options.documentText ??
      contextText,
  });
  insertedRange.insertContentControl = vi.fn(() => cc);
  insertedRange.insertComment = vi.fn();

  const anchorRange = createRange({
    text: anchorText,
    parentCC: options.anchorRangeParentCC,
    paragraphText: contextText,
    insertTextImpl: (text: string) => {
      if (options.insertError) {
        throw options.insertError;
      }
      if (!insertedCurrentText) {
        insertedRange.text = text;
        insertedRange.getReviewedText = vi.fn((version: "Current" | "Original") =>
          createClientResult(version === "Original" ? "" : text)
        );
      }
      return insertedRange;
    },
  });

  const anchorSearchSequence = options.anchorSearchSequence ?? [[anchorRange]];
  const wrapperSearchSequence = [
    [anchorRange],
    ...(options.insertedRange?.searchSequence ?? [[insertedRange]]),
  ];
  operationalWrapperRange.search = createSearchMock(wrapperSearchSequence);
  const bodyRange = createRange({
    text: contextText,
    searchSequence: anchorSearchSequence,
    parentCC: options.contextRangeParentCC,
  });

  const paragraphSearchSeq = options.paragraphSearchSequence ?? anchorSearchSequence;
  const paragraphSearchMock = createSearchMock(paragraphSearchSeq);

  (bodyRange.paragraphs as any).getFirst().getRange("Whole").search = paragraphSearchMock;

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

  const wordGlobal = globalThis as unknown as {
    Word?: {
      ChangeTrackingMode: Record<string, string>;
      InsertLocation: Record<string, string>;
      run: ReturnType<typeof vi.fn>;
    };
  };
  wordGlobal.Word = {
    ChangeTrackingMode: {
      off: "off",
      trackAll: "trackAll",
      trackMine: "trackMine",
    },
    InsertLocation: {
      replace: "Replace",
    },
    run: vi.fn(async (callback: (ctx: typeof context) => unknown) => callback(context)),
  };

  anchorRange.insertContentControl = vi.fn(() =>
    options.useOperationalWrapper === false ? cc : operationalWrapper
  );
  bodyRange.insertContentControl = vi.fn(() => operationalWrapper);

  const testContext: ApplyCommandTestContext = {
    context,
    bodyRange,
    anchorRange,
    insertedRange,
    cc,
    operationalWrapper,
    operationalWrapperRange,
  };

  options.setupParagraphSearch?.(
    testContext,
    bodyRange,
    { current: null },
    options.anchorRangeRef ?? { current: null }
  );

  return testContext;
}
