import type { Suggestion } from "../../domain/suggestion/Suggestion.types";

const OPERATIONAL_WRAPPER_TITLE_PREFIX = "stylistic-meta-v2:";
const OPERATIONAL_WRAPPER_TAG_PREFIX = "stylistic-operational-wrapper:";

type MockTrackedChange = {
  id?: string;
  text?: string;
  type?: string;
  accept?: ReturnType<typeof vi.fn>;
  reject?: ReturnType<typeof vi.fn>;
  getRange?: ReturnType<typeof vi.fn>;
};

type MockRangeWithTrackedChanges = {
  compareLocationWith: ReturnType<typeof vi.fn>;
  getTrackedChanges: ReturnType<typeof vi.fn>;
};

type MockTrackedChangeCollection = {
  items: MockTrackedChange[];
  load: ReturnType<typeof vi.fn>;
  acceptAll: ReturnType<typeof vi.fn>;
  rejectAll: ReturnType<typeof vi.fn>;
};

type MockComment = {
  authorName?: string;
  content?: string;
  getRange: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
};

type MockCommentRange = {
  compareLocationWith: ReturnType<typeof vi.fn>;
  getTrackedChanges: ReturnType<typeof vi.fn>;
};

type ResolveSuggestionContext = {
  document: {
    contentControls: {
      getByTag: ReturnType<typeof vi.fn>;
      load: ReturnType<typeof vi.fn>;
      items: Array<{ tag: string }>;
    };
    load: ReturnType<typeof vi.fn>;
    changeTrackingMode: string;
    body: {
      search: ReturnType<typeof vi.fn>;
      load: ReturnType<typeof vi.fn>;
      text: string;
      getComments: ReturnType<typeof vi.fn>;
      getTrackedChanges: ReturnType<typeof vi.fn>;
    };
  };
  sync: ReturnType<typeof vi.fn>;
  _ccsCollection: { items: unknown[]; load: ReturnType<typeof vi.fn> };
  _commentsCollection: { items: MockComment[]; load: ReturnType<typeof vi.fn> };
  _bodyTCCollection: { items: MockTrackedChange[]; load: ReturnType<typeof vi.fn> };
  _rangeTCCollection: { items: MockTrackedChange[]; load: ReturnType<typeof vi.fn> };
  _commentRangeTCCollections: Array<{ items: MockTrackedChange[]; load: ReturnType<typeof vi.fn> }>;
  _ccItems: Array<{
    title: string;
    tag: string;
    load: ReturnType<typeof vi.fn>;
    getTrackedChanges: ReturnType<typeof vi.fn>;
    getRange: ReturnType<typeof vi.fn>;
    delete: ReturnType<typeof vi.fn>;
  }>;
  _cc: {
    title: string;
    tag: string;
    load: ReturnType<typeof vi.fn>;
    getTrackedChanges: ReturnType<typeof vi.fn>;
    getRange: ReturnType<typeof vi.fn>;
    delete: ReturnType<typeof vi.fn>;
  };
};

type MockWordGlobal = {
  run: ReturnType<typeof vi.fn>;
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

/** Serializes a default operational-wrapper title payload for replace-suggestion tests. */
export function makeOperationalWrapperTitle(options: {
  suggestionId?: string;
  insertedTag?: string;
  deletedValue?: string;
  anchorValue?: string;
  groupId?: string;
  groupIndex?: number;
  groupSize?: number;
  overrides?: Record<string, unknown>;
} = {}): string {
  return `${OPERATIONAL_WRAPPER_TITLE_PREFIX}${JSON.stringify({
    suggestionId: options.suggestionId ?? "s-1",
    version: "operational-wrapper-v1",
    insertedSideRef: {
      kind: "content-control",
      role: "inserted-side",
      value: options.insertedTag ?? "stylistic:track-change:s-1",
    },
    deletedSideRef: {
      kind: "anchor",
      role: "deleted-side",
      value: options.deletedValue ?? "texto original",
    },
    anchorRef: {
      kind: "anchor",
      role: "operational-anchor",
      value: options.anchorValue ?? "Contexto con texto original.",
    },
    groupId: options.groupId ?? options.suggestionId ?? "s-1",
    groupIndex: options.groupIndex ?? 0,
    groupSize: options.groupSize ?? 1,
    ...options.overrides,
  })}`;
}

/** Builds the external operational-wrapper tag used by live resolution. */
export function makeOperationalWrapperTag(suggestionId = "s-1"): string {
  return `${OPERATIONAL_WRAPPER_TAG_PREFIX}${suggestionId}`;
}

/**
 * Installs a `Word.run` mock that executes the callback with the provided context.
 */
export function installWordWithContext<TContext>(context: TContext) {
  const run = vi.fn(async <T>(
    callback: (ctx: TContext) => Promise<T> | T
  ) => callback(context));
  const wordGlobal = globalThis as unknown as {
    Word?: MockWordGlobal & { ChangeTrackingMode: Record<string, string> };
  };
  wordGlobal.Word = {
    run,
    ChangeTrackingMode: {
      off: "off",
      trackAll: "trackAll",
      trackMine: "trackMine",
    },
  };
  return run;
}

/**
 * Installs a rejecting `Word.run` mock for sad-path adapter tests.
 */
export function installRejectingWord(error: Error) {
  const run = vi.fn(async () => {
    throw error;
  });
  const wordGlobal = globalThis as unknown as {
    Word?: MockWordGlobal & { ChangeTrackingMode: Record<string, string> };
  };
  wordGlobal.Word = {
    run,
    ChangeTrackingMode: {
      off: "off",
      trackAll: "trackAll",
      trackMine: "trackMine",
    },
  };
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
  ccTitle,
  ccTag = makeOperationalWrapperTag("s-1"),
  ccItems,
  spanTCItems = [] as MockTrackedChange[],
  rangeTCItems = [] as MockTrackedChange[],
  bodyTCItems = [] as MockTrackedChange[],
  bodyTCRelations = [] as string[],
  comments = [] as MockComment[],
  commentRangeTCItems = [] as MockTrackedChange[][],
  deletedSideText,
  deletedSideRangeTCItems = [] as MockTrackedChange[],
  operationalAnchorText,
  operationalAnchorRangeTCItems = [] as MockTrackedChange[],
}: {
  ccFound?: boolean;
  ccTitle?: string;
  ccTag?: string;
  ccItems?: Array<{
    title?: string;
    tag?: string;
    spanTCItems?: MockTrackedChange[];
    rangeTCItems?: MockTrackedChange[];
    rangeRelationWithNext?: string;
  }>;
  spanTCItems?: MockTrackedChange[];
  rangeTCItems?: MockTrackedChange[];
  bodyTCItems?: MockTrackedChange[];
  bodyTCRelations?: string[];
  comments?: MockComment[];
  commentRangeTCItems?: MockTrackedChange[][];
  deletedSideText?: string;
  deletedSideRangeTCItems?: MockTrackedChange[];
  operationalAnchorText?: string;
  operationalAnchorRangeTCItems?: MockTrackedChange[];
}): ResolveSuggestionContext {
  const effectiveSpanTCItems =
    spanTCItems.length > 0 ? spanTCItems : rangeTCItems;
  const effectiveRangeTCItems =
    rangeTCItems.length > 0 ? rangeTCItems : effectiveSpanTCItems;
  const mutableSpanTCItems = [...effectiveSpanTCItems];
  const mutableRangeTCItems = [...effectiveRangeTCItems];
  const mutableBodyTCItems = [...bodyTCItems];
  const mutableDeletedSideRangeTCItems = [...deletedSideRangeTCItems];
  const mutableOperationalAnchorRangeTCItems = [...operationalAnchorRangeTCItems];
  const mutableCommentRangeTCItems = commentRangeTCItems.map((items) => [...items]);

  const trackedChangeCollections = [
    mutableSpanTCItems,
    mutableRangeTCItems,
    mutableBodyTCItems,
    mutableDeletedSideRangeTCItems,
    mutableOperationalAnchorRangeTCItems,
    ...mutableCommentRangeTCItems,
  ];

  const removeTrackedChangesBySemanticSide = (trackedChange: MockTrackedChange) => {
    const trackedChangeType = trackedChange.type;
    if (!trackedChangeType) {
      return;
    }

    for (const collection of trackedChangeCollections) {
      for (let index = collection.length - 1; index >= 0; index -= 1) {
        if (collection[index]?.type === trackedChangeType) {
          collection.splice(index, 1);
        }
      }
    }
  };

  const wrappedTrackedChanges = new WeakSet<object>();
  const wrapTrackedChangeMutation = (trackedChange: MockTrackedChange) => {
    if (wrappedTrackedChanges.has(trackedChange)) {
      return trackedChange;
    }

    const originalAccept = trackedChange.accept;
    const originalReject = trackedChange.reject;

    if (originalAccept) {
      trackedChange.accept = vi.fn(() => {
        const result = (originalAccept as () => unknown)();
        removeTrackedChangesBySemanticSide(trackedChange);
        return result;
      });
    }

    if (originalReject) {
      trackedChange.reject = vi.fn(() => {
        const result = (originalReject as () => unknown)();
        removeTrackedChangesBySemanticSide(trackedChange);
        return result;
      });
    }

    wrappedTrackedChanges.add(trackedChange);
    return trackedChange;
  };

  for (const collection of trackedChangeCollections) {
    for (const trackedChange of collection) {
      wrapTrackedChangeMutation(trackedChange);
    }
  }

  const ccTagParts = ccTag.split(":");
  const inferredSuggestionId = ccTagParts[ccTagParts.length - 1] ?? "s-1";
  const buildTrackedChangeCollection = (
    items: MockTrackedChange[],
  ): MockTrackedChangeCollection => ({
    items,
    load: vi.fn(),
    acceptAll: vi.fn(() => {
      for (const trackedChange of [...items]) {
        const accept = trackedChange.accept as (() => unknown) | undefined;
        accept?.();
      }
    }),
    rejectAll: vi.fn(() => {
      for (const trackedChange of [...items]) {
        const reject = trackedChange.reject as (() => unknown) | undefined;
        reject?.();
      }
    }),
  });

  const rangeTCCollection = buildTrackedChangeCollection(mutableRangeTCItems);
  const bodyTCCollection = buildTrackedChangeCollection(mutableBodyTCItems);
  const operationalAnchorRangeTCCollection = buildTrackedChangeCollection(
    mutableOperationalAnchorRangeTCItems,
  );
  const deletedSideRangeTCCollection = buildTrackedChangeCollection(
    mutableDeletedSideRangeTCItems,
  );

  const buildCc = (options?: {
    title?: string;
    tag?: string;
    spanTCItems?: MockTrackedChange[];
    rangeTCItems?: MockTrackedChange[];
    rangeRelationWithNext?: string;
  }) => {
    const thisTag = options?.tag ?? ccTag;
    const thisSpanTCItems = options?.spanTCItems ?? mutableSpanTCItems;
    const thisRangeTCItems = options?.rangeTCItems ?? mutableRangeTCItems;

    if (!trackedChangeCollections.includes(thisSpanTCItems)) {
      trackedChangeCollections.push(thisSpanTCItems);
    }

    if (!trackedChangeCollections.includes(thisRangeTCItems)) {
      trackedChangeCollections.push(thisRangeTCItems);
    }

    for (const trackedChange of thisSpanTCItems) {
      wrapTrackedChangeMutation(trackedChange);
    }

    for (const trackedChange of thisRangeTCItems) {
      wrapTrackedChangeMutation(trackedChange);
    }

    const thisSpanTCCollection = buildTrackedChangeCollection(thisSpanTCItems);
    const thisRangeTCCollection = buildTrackedChangeCollection(thisRangeTCItems);
    const thisCcRange: MockRangeWithTrackedChanges = {
      compareLocationWith: vi.fn(() => ({
        value: options?.rangeRelationWithNext ?? "AdjacentBefore",
      })),
      getTrackedChanges: vi.fn(() => thisRangeTCCollection),
    };

    return {
      title:
        options?.title ??
        ccTitle ??
        (thisTag.startsWith(OPERATIONAL_WRAPPER_TAG_PREFIX)
          ? makeOperationalWrapperTitle({
              suggestionId: inferredSuggestionId,
              insertedTag: `stylistic:track-change:${inferredSuggestionId}`,
            })
          : "texto original"),
      tag: thisTag,
      load: vi.fn(),
      getTrackedChanges: vi.fn(() => thisSpanTCCollection),
      getRange: vi.fn(() => thisCcRange),
      delete: vi.fn(),
    };
  };

  const cc = buildCc();
  const allCcItems = ccItems?.map((item) => buildCc(item)) ?? (ccFound ? [cc] : []);

  const bodyTrackedChangeRanges = bodyTCItems.map((_, index) => ({
    compareLocationWith: vi.fn(() => ({
      value: bodyTCRelations[index] ?? "Disjoint",
    })),
  }));

  for (let index = 0; index < bodyTCItems.length; index += 1) {
    const trackedChange = bodyTCItems[index];
    trackedChange.getRange = vi.fn(() => bodyTrackedChangeRanges[index]);
  }

  const commentRangeTCCollections = comments.map((_, index) =>
    buildTrackedChangeCollection(mutableCommentRangeTCItems[index] ?? []),
  );

  comments.forEach((comment, index) => {
    const originalRange = {
      compareLocationWith: vi.fn(() => ({ value: "Equal" })),
    } satisfies { compareLocationWith: ReturnType<typeof vi.fn> };
    const enrichedRange: MockCommentRange = {
      compareLocationWith: originalRange.compareLocationWith,
      getTrackedChanges: vi.fn(() => commentRangeTCCollections[index]),
    };
    comment.getRange = vi.fn(() => enrichedRange);
  });

  const ccsCollection = {
    items: allCcItems,
    load: vi.fn(),
  };

  const documentContentControls = {
    getByTag: vi.fn(() => ccsCollection),
    load: vi.fn(),
    items: allCcItems,
  };

  const commentsCollection = { items: comments, load: vi.fn() };

  const operationalAnchorRange = {
    text: operationalAnchorText ?? "",
    load: vi.fn(),
    search: vi.fn(() => ({ items: [], load: vi.fn() })),
    getTrackedChanges: vi.fn(() => operationalAnchorRangeTCCollection),
  };

  const deletedSideRange = {
    text: deletedSideText ?? "",
    load: vi.fn(),
    search: vi.fn(() => ({ items: [], load: vi.fn() })),
    getTrackedChanges: vi.fn(() => deletedSideRangeTCCollection),
  };

  const bodySearch = vi.fn((searchText: string) => ({
    items: (() => {
      if (deletedSideText && searchText === deletedSideText) {
        return [deletedSideRange];
      }

      if (operationalAnchorText && searchText === operationalAnchorText) {
        return [operationalAnchorRange];
      }

      return [];
    })(),
    load: vi.fn(),
  }));

  return {
    document: {
      contentControls: documentContentControls,
      load: vi.fn(),
      changeTrackingMode: "trackAll",
      body: {
        search: bodySearch,
        load: vi.fn(),
        text: operationalAnchorText ?? "",
        getComments: vi.fn(() => commentsCollection),
        getTrackedChanges: vi.fn(() => bodyTCCollection),
      },
    },
    sync: vi.fn().mockResolvedValue(undefined),
    _ccsCollection: ccsCollection,
    _commentsCollection: commentsCollection,
    _bodyTCCollection: bodyTCCollection,
    _rangeTCCollection: rangeTCCollection,
    _commentRangeTCCollections: commentRangeTCCollections,
    _ccItems: allCcItems,
    _cc: cc,
  };
}
