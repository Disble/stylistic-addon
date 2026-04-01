import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { Suggestion } from "../../domain/types";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";

type ParentCC = {
  tag: string;
  isNullObject: boolean;
  load: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
};

type ApplyTestContext = {
  context: {
    document: {
      body: {
        search: ReturnType<typeof vi.fn>;
        load: ReturnType<typeof vi.fn>;
        text: string;
      };
      load: ReturnType<typeof vi.fn>;
      changeTrackingMode: string;
    };
    sync: ReturnType<typeof vi.fn>;
  };
  range: {
    insertText: ReturnType<typeof vi.fn>;
    delete: ReturnType<typeof vi.fn>;
    insertComment: ReturnType<typeof vi.fn>;
    parentContentControlOrNullObject: ParentCC;
  };
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

type SearchRange = {
  insertText: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
  insertComment: ReturnType<typeof vi.fn>;
  insertContentControl: ReturnType<typeof vi.fn>;
  parentContentControlOrNullObject: ParentCC;
};

type SearchResults = {
  items: SearchRange[];
  load: ReturnType<typeof vi.fn>;
};

function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  return {
    id: "s1",
    originalText: "texto original",
    suggestedText: "texto sugerido",
    justification: "Mejora la claridad",
    category: "Estilo",
    severity: "medium",
    type: "track-change",
    ...overrides,
  };
}

function installWordContext(
  options: {
    resultsItems?: SearchRange[];
    searchResultsSequence?: SearchRange[][];
    documentText?: string;
    initialTrackingMode?: string;
    insertError?: Error;
    onSync?: (count: number) => void | Promise<void>;
    parentCC?: Partial<ParentCC>;
  } = {},
): ApplyTestContext {
  const {
    resultsItems,
    searchResultsSequence,
    documentText = "",
    initialTrackingMode = "off",
    insertError,
    onSync,
    parentCC: parentCCOverride,
  } = options;

  const parentCC: ParentCC = {
    tag: "",
    isNullObject: true,
    load: vi.fn(),
    delete: vi.fn(),
    ...parentCCOverride,
  };

  const cc = { tag: "", appearance: "", cannotDelete: true };
  const insertedRange = {
    insertContentControl: vi.fn(() => cc),
    insertComment: vi.fn(),
  };

  const range: SearchRange = {
    insertText: vi.fn(() => {
      if (insertError) {
        throw insertError;
      }
      return insertedRange;
    }),
    delete: vi.fn(),
    insertComment: vi.fn(),
    insertContentControl: vi.fn(() => cc),
    parentContentControlOrNullObject: parentCC,
  };

  const searchCalls: SearchResults[] = [];

  let syncCount = 0;
  const context = {
    document: {
      body: {
        search: vi.fn(() => {
          const nextItems =
            searchResultsSequence?.[searchCalls.length] ??
            resultsItems ?? [range];
          const results = {
            items: nextItems,
            load: vi.fn(),
          };
          searchCalls.push(results);
          return results;
        }),
        load: vi.fn(),
        text: documentText,
      },
      load: vi.fn(),
      changeTrackingMode: initialTrackingMode,
    },
    sync: vi.fn(async () => {
      syncCount += 1;
      await onSync?.(syncCount);
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

  return { context, range, insertedRange, cc };
}

describe("ApplySuggestionCommand", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
    vi.spyOn(console, "warn").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  // ─── SC-TC-01: Replace path ───────────────────────────────────────────────

  it("SC-TC-01: sets trackAll, calls insertText(replace), restores mode, wraps in CC", async () => {
    const { context, insertedRange, cc } = installWordContext({
      initialTrackingMode: "off",
    });

    const command = new ApplySuggestionCommand(makeSuggestion());
    const result = await command.execute();

    // Result
    expect(result).toEqual({ success: true, commandId: "s1" });

    // changeTrackingMode was set to trackAll BEFORE insertion
    // (verified by the sequence: it must have been trackAll at insertion time,
    // and restored to "off" after)
    expect(context.document.changeTrackingMode).toBe("off");

    // insertText called with (suggestedText, replace)
    const { range } = installWordContext(); // just for type ref — we use context above
    // We check via the search result range — but we need the actual range used.
    // The context.document.body.search returns range items, so get them via mock:
    const searchResult = context.document.body.search.mock.results[0]
      ?.value as SearchResults;
    const usedRange = searchResult.items[0];
    expect(usedRange.insertText).toHaveBeenCalledWith(
      "texto sugerido",
      "Replace",
    );

    // CC was tagged correctly
    expect(cc.tag).toBe("stylistic:track-change:s1");
    expect(cc.appearance).toBe("Hidden");
    expect(cc.cannotDelete).toBe(false);

    // insertedRange.insertContentControl was called
    expect(insertedRange.insertContentControl).toHaveBeenCalled();
  });

  it("SC-TC-01b: changeTrackingMode set to trackAll before insertText and restored to prior value in finally", async () => {
    const modeSequence: string[] = [];

    const { context } = installWordContext({
      initialTrackingMode: "off",
      onSync: async () => {
        modeSequence.push(context.document.changeTrackingMode);
      },
    });

    const command = new ApplySuggestionCommand(makeSuggestion());
    await command.execute();

    // At some sync point, mode was "trackAll"
    expect(modeSequence).toContain("trackAll");
    // Final value is restored to "off"
    expect(context.document.changeTrackingMode).toBe("off");
  });

  // ─── SC-TC-02: Delete path ────────────────────────────────────────────────

  it("SC-TC-02: delete path — insertText('', replace) called when suggestedText is empty", async () => {
    const { context } = installWordContext({ initialTrackingMode: "off" });

    const command = new ApplySuggestionCommand(
      makeSuggestion({ suggestedText: "" }),
    );
    const result = await command.execute();

    expect(result).toEqual({ success: true, commandId: "s1" });

    const searchResult = context.document.body.search.mock.results[0]
      ?.value as SearchResults;
    const usedRange = searchResult.items[0];
    expect(usedRange.insertText).toHaveBeenCalledWith("", "Replace");
    // changeTrackingMode restored
    expect(context.document.changeTrackingMode).toBe("off");
  });

  // ─── SC-TC-03: Insert without anchor (unsupported) ────────────────────────

  it("SC-TC-03: returns { success: false } for insert-only suggestions without touching document", async () => {
    const { context } = installWordContext();

    const command = new ApplySuggestionCommand(
      makeSuggestion({ originalText: "", suggestedText: "texto sugerido" }),
    );
    const result = await command.execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Insert-only suggestions require anchor text",
    });
    expect(context.document.body.search).not.toHaveBeenCalled();
    expect(context.document.load).not.toHaveBeenCalled();
  });

  // ─── SC-TC-04: Original text not found ───────────────────────────────────

  it("SC-TC-04: returns { success: false } when original text not found — changeTrackingMode NOT modified", async () => {
    const { context } = installWordContext({
      resultsItems: [],
      initialTrackingMode: "off",
    });

    const command = new ApplySuggestionCommand(makeSuggestion());
    const result = await command.execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Texto original no encontrado",
    });
    // changeTrackingMode must not have been modified
    expect(context.document.changeTrackingMode).toBe("off");
    expect(context.document.load).not.toHaveBeenCalled();
  });

  // ─── SC-TC-05: changeTrackingMode always restored even on error ───────────

  it("SC-TC-05: restores changeTrackingMode to 'off' even when insertText throws", async () => {
    const insertionError = new Error("insert failed");
    const { context } = installWordContext({
      insertError: insertionError,
      initialTrackingMode: "off",
    });

    const command = new ApplySuggestionCommand(makeSuggestion());
    const result = await command.execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "insert failed",
    });
    expect(context.document.changeTrackingMode).toBe("off");
  });

  // ─── SC-CO-01: Comment-only happy path ───────────────────────────────────

  it("SC-CO-01: comment-only — calls range.insertComment, NO changeTrackingMode change, CC tagged", async () => {
    const { context, cc } = installWordContext({
      initialTrackingMode: "trackAll",
    });

    const command = new ApplySuggestionCommand(
      makeSuggestion({
        id: "chunk1-5",
        type: "comment-only",
        suggestedText: undefined,
        originalText: "texto original",
        category: "Estilo",
        justification: "Mejora la claridad",
      }),
    );
    const result = await command.execute();

    expect(result).toEqual({ success: true, commandId: "chunk1-5" });

    // Comment inserted on the range
    const searchResult = context.document.body.search.mock.results[0]
      ?.value as SearchResults;
    const usedRange = searchResult.items[0];
    expect(usedRange.insertComment).toHaveBeenCalledWith(
      "[Estilo]\nMejora la claridad",
    );

    // CC tagged correctly
    expect(cc.tag).toBe("stylistic:comment-only:chunk1-5");
    expect(cc.appearance).toBe("Hidden");

    // changeTrackingMode NOT touched
    expect(context.document.changeTrackingMode).toBe("trackAll");
    expect(context.document.load).not.toHaveBeenCalled();
  });

  // ─── SC-CO-02: Comment-only text not found ───────────────────────────────

  it("SC-CO-02: comment-only — returns { success: false } when text not found", async () => {
    const { context } = installWordContext({ resultsItems: [] });

    const command = new ApplySuggestionCommand(
      makeSuggestion({ type: "comment-only", suggestedText: undefined }),
    );
    const result = await command.execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Texto original no encontrado",
    });
    expect(context.document.body.search).toHaveBeenCalledWith("texto original", {
      matchCase: true,
      matchWholeWord: false,
    });
  });

  // ─── SC-CO-03: changeTrackingMode never touched in comment-only ───────────

  it("SC-CO-03: changeTrackingMode never read or set in comment-only path", async () => {
    const { context } = installWordContext({
      initialTrackingMode: "trackAll",
    });

    const command = new ApplySuggestionCommand(
      makeSuggestion({ type: "comment-only", suggestedText: undefined }),
    );
    await command.execute();

    // document.load must NOT have been called (loading changeTrackingMode)
    expect(context.document.load).not.toHaveBeenCalled();
    // Value unchanged
    expect(context.document.changeTrackingMode).toBe("trackAll");
  });

  // ─── Existing search/fallback behaviors ──────────────────────────────────

  it("returns a failed result when the original text is not found", async () => {
    const { context } = installWordContext({ resultsItems: [] });

    const command = new ApplySuggestionCommand(makeSuggestion());
    const result = await command.execute();

    const results = context.document.body.search.mock.results[0]
      ?.value as SearchResults;

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Texto original no encontrado",
    });
    expect(context.document.body.search).toHaveBeenCalledWith(
      "texto original",
      {
        matchCase: true,
        matchWholeWord: false,
      },
    );
    expect(results.load).toHaveBeenCalledWith("items");
    expect(context.document.load).not.toHaveBeenCalled();
  });

  it("falls back to the exact document slice when backend whitespace differs from Word text", async () => {
    const originalText =
      "Por eso ninguno de nosotros sabe nada en realidad.»Así que no creo que XXXX esté tratando de ocultarte información.";
    const documentText =
      "Por eso ninguno de nosotros sabe nada en realidad.\r»Así que no creo que XXXX esté tratando de ocultarte información.";

    const { context, range } = installWordContext({
      documentText,
    });
    const fallbackRange: SearchRange = {
      insertText: range.insertText,
      delete: vi.fn(),
      insertComment: vi.fn(),
      insertContentControl: vi.fn(() => ({ tag: "", appearance: "", cannotDelete: true })),
      parentContentControlOrNullObject: {
        tag: "",
        isNullObject: true,
        load: vi.fn(),
        delete: vi.fn(),
      },
    };
    context.document.body.search.mockReset();
    context.document.body.search
      .mockImplementationOnce(() => ({ items: [], load: vi.fn() })) // attempt 1: exact
      .mockImplementationOnce(() => ({ items: [], load: vi.fn() })) // attempt 1.5: ignorePunct+ignoreSpace
      .mockImplementationOnce(() => ({
        items: [fallbackRange],
        load: vi.fn(),
      })); // attempt 2: whitespace fallback

    const command = new ApplySuggestionCommand(
      makeSuggestion({ originalText }),
    );
    const result = await command.execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(context.document.body.search).toHaveBeenNthCalledWith(
      1,
      originalText,
      {
        matchCase: true,
        matchWholeWord: false,
      },
    );
    expect(context.document.body.load).toHaveBeenCalledWith("text");
    expect(context.document.body.search).toHaveBeenNthCalledWith(
      3,
      documentText,
      {
        matchCase: true,
        matchWholeWord: false,
      },
    );
  });

  it("falls back when backend flattens paragraph breaks into spaces", async () => {
    const originalText =
      "los usuarios áuricos. —¿Eh?—Los humanos tampoco saben mucho sobre los humanos, ¿verdad? —dijo como si fuera una obviedad—.";
    const documentText =
      "los usuarios áuricos. \r—¿Eh?\r—Los humanos tampoco saben mucho sobre los humanos, ¿verdad? —dijo como si fuera una obviedad—.";

    const { context, range } = installWordContext({ documentText });
    const fallbackRange: SearchRange = {
      insertText: range.insertText,
      delete: vi.fn(),
      insertComment: vi.fn(),
      insertContentControl: vi.fn(() => ({ tag: "", appearance: "", cannotDelete: true })),
      parentContentControlOrNullObject: {
        tag: "",
        isNullObject: true,
        load: vi.fn(),
        delete: vi.fn(),
      },
    };
    context.document.body.search.mockReset();
    context.document.body.search
      .mockImplementationOnce(() => ({ items: [], load: vi.fn() })) // attempt 1: exact
      .mockImplementationOnce(() => ({ items: [], load: vi.fn() })) // attempt 1.5: ignorePunct+ignoreSpace
      .mockImplementationOnce(() => ({
        items: [fallbackRange],
        load: vi.fn(),
      })); // attempt 2: whitespace fallback

    const command = new ApplySuggestionCommand(
      makeSuggestion({ originalText }),
    );
    const result = await command.execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(context.document.body.load).toHaveBeenCalledWith("text");
    expect(context.document.body.search).toHaveBeenNthCalledWith(
      3,
      documentText,
      {
        matchCase: true,
        matchWholeWord: false,
      },
    );
  });

  it("classifies replace suggestions and calls insertText with suggestedText", async () => {
    const { context } = installWordContext();

    const command = new ApplySuggestionCommand(
      makeSuggestion({
        originalText: "texto original",
        suggestedText: "texto sugerido",
      }),
    );
    await command.execute();

    expect(context.document.body.search).toHaveBeenCalledWith(
      "texto original",
      {
        matchCase: true,
        matchWholeWord: false,
      },
    );
    const searchResult = context.document.body.search.mock.results[0]
      ?.value as SearchResults;
    const usedRange = searchResult.items[0];
    expect(usedRange.insertText).toHaveBeenCalledWith("texto sugerido", "Replace");
  });

  it("classifies delete suggestions and calls insertText with empty string", async () => {
    const { context } = installWordContext();

    const command = new ApplySuggestionCommand(
      makeSuggestion({ originalText: "texto original", suggestedText: "" }),
    );
    await command.execute();

    expect(context.document.body.search).toHaveBeenCalledWith(
      "texto original",
      {
        matchCase: true,
        matchWholeWord: false,
      },
    );
    const searchResult = context.document.body.search.mock.results[0]
      ?.value as SearchResults;
    const usedRange = searchResult.items[0];
    expect(usedRange.insertText).toHaveBeenCalledWith("", "Replace");
  });

  it("returns a failed result for insert-only suggestions without searching for empty text", async () => {
    const { context } = installWordContext();

    const command = new ApplySuggestionCommand(
      makeSuggestion({ originalText: "", suggestedText: "texto sugerido" }),
    );

    const result = await command.execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Insert-only suggestions require anchor text",
    });
    expect(context.document.body.search).not.toHaveBeenCalled();
    expect(context.document.load).not.toHaveBeenCalled();
  });

  it("assigns a namespaced CC tag with the suggestion type for track-change suggestions", async () => {
    const { cc } = installWordContext();

    const command = new ApplySuggestionCommand(
      makeSuggestion({ id: "chunk0-3" }),
    );
    await command.execute();

    expect(cc.tag).toBe("stylistic:track-change:chunk0-3");
  });

  describe("already-covered detection (CC nesting prevention)", () => {
    function makeCleanRange(): SearchRange {
      const cleanCC = { tag: "", appearance: "", cannotDelete: true };
      const cleanInsertedRange = {
        insertContentControl: vi.fn(() => cleanCC),
        insertComment: vi.fn(),
      };
      return {
        insertText: vi.fn(() => cleanInsertedRange),
        delete: vi.fn(),
        insertComment: vi.fn(),
        insertContentControl: vi.fn(() => cleanCC),
        parentContentControlOrNullObject: {
          tag: "",
          isNullObject: true,
          load: vi.fn(),
          delete: vi.fn(),
        },
      };
    }

    it("deletes the existing Stylistic CC, re-searches, and inserts for track-change", async () => {
      const coveredCC: ParentCC = {
        tag: "stylistic:track-change:chunk0-0",
        isNullObject: false,
        load: vi.fn(),
        delete: vi.fn(),
      };
      const coveredRange: SearchRange = {
        insertText: vi.fn(),
        delete: vi.fn(),
        insertComment: vi.fn(),
        insertContentControl: vi.fn(),
        parentContentControlOrNullObject: coveredCC,
      };
      const freshRange = makeCleanRange();
      const { context } = installWordContext({
        searchResultsSequence: [[coveredRange], [freshRange]],
      });

      const command = new ApplySuggestionCommand(makeSuggestion());
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(coveredCC.delete).toHaveBeenCalledWith(true);
      expect(context.document.body.search).toHaveBeenCalledTimes(2);
      expect(coveredRange.insertText).not.toHaveBeenCalled();
      expect(freshRange.insertText).toHaveBeenCalled();
    });

    it("deletes the existing legacy chunk CC, re-searches, and inserts", async () => {
      const coveredCC: ParentCC = {
        tag: "chunk0-0",
        isNullObject: false,
        load: vi.fn(),
        delete: vi.fn(),
      };
      const coveredRange: SearchRange = {
        insertText: vi.fn(),
        delete: vi.fn(),
        insertComment: vi.fn(),
        insertContentControl: vi.fn(),
        parentContentControlOrNullObject: coveredCC,
      };
      const freshRange = makeCleanRange();
      const { context } = installWordContext({
        searchResultsSequence: [[coveredRange], [freshRange]],
      });

      const command = new ApplySuggestionCommand(makeSuggestion());
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(coveredCC.delete).toHaveBeenCalledWith(true);
      expect(context.document.body.search).toHaveBeenCalledTimes(2);
      expect(coveredRange.insertText).not.toHaveBeenCalled();
      expect(freshRange.insertText).toHaveBeenCalled();
    });

    it("proceeds normally when range has no parent CC (isNullObject: true)", async () => {
      const { range } = installWordContext({
        parentCC: { isNullObject: true, tag: "" },
      });

      const command = new ApplySuggestionCommand(makeSuggestion());
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(
        range.parentContentControlOrNullObject.delete,
      ).not.toHaveBeenCalled();
      expect(range.insertText).toHaveBeenCalled();
    });

    it("returns { success: false } when re-search after CC deletion finds nothing", async () => {
      const coveredCC: ParentCC = {
        tag: "stylistic:track-change:chunk0-0",
        isNullObject: false,
        load: vi.fn(),
        delete: vi.fn(),
      };
      const coveredRange: SearchRange = {
        insertText: vi.fn(),
        delete: vi.fn(),
        insertComment: vi.fn(),
        insertContentControl: vi.fn(),
        parentContentControlOrNullObject: coveredCC,
      };
      installWordContext({
        searchResultsSequence: [[coveredRange], []],
      });

      const command = new ApplySuggestionCommand(makeSuggestion());
      const result = await command.execute();

      expect(result).toEqual({
        success: false,
        commandId: "s1",
        error: "Texto no encontrado tras eliminar CC existente",
      });
      // Must use keepContent: true — never delete document content on failure
      expect(coveredCC.delete).toHaveBeenCalledWith(true);
    });

    it("deletes the existing Stylistic CC, re-searches, and inserts for comment-only", async () => {
      const coveredCC: ParentCC = {
        tag: "stylistic:comment-only:chunk1-5",
        isNullObject: false,
        load: vi.fn(),
        delete: vi.fn(),
      };
      const coveredRange: SearchRange = {
        insertText: vi.fn(),
        delete: vi.fn(),
        insertComment: vi.fn(),
        insertContentControl: vi.fn(),
        parentContentControlOrNullObject: coveredCC,
      };
      const freshRange = makeCleanRange();
      const { context } = installWordContext({
        searchResultsSequence: [[coveredRange], [freshRange]],
      });

      const command = new ApplySuggestionCommand(
        makeSuggestion({ type: "comment-only", suggestedText: undefined }),
      );
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(coveredCC.delete).toHaveBeenCalledWith(true);
      expect(context.document.body.search).toHaveBeenCalledTimes(2);
      expect(coveredRange.insertText).not.toHaveBeenCalled();
      // freshRange gets insertComment called (comment-only path)
      expect(freshRange.insertComment).toHaveBeenCalled();
    });

    it("stateful mock: text survives CC removal — self-enforcing without keepContent assertion", async () => {
      // Behavioral mock: search() closes over textDeleted.
      // delete(false) → textDeleted=true → 2nd search returns [] → success:false → test FAILS
      // delete(true)  → textDeleted=false → 2nd search returns freshRange → inserted → PASSES
      let textDeleted = false;
      let searchCalls = 0;

      const coveredCC: ParentCC = {
        tag: "stylistic:track-change:chunk0-0",
        isNullObject: false,
        load: vi.fn(),
        delete: vi.fn((keepContent: boolean) => {
          textDeleted = !keepContent;
        }),
      };
      const coveredRange: SearchRange = {
        insertText: vi.fn(),
        delete: vi.fn(),
        insertComment: vi.fn(),
        insertContentControl: vi.fn(),
        parentContentControlOrNullObject: coveredCC,
      };
      const freshRange = makeCleanRange();

      const context = {
        document: {
          body: {
            search: vi.fn(() => {
              const call = searchCalls++;
              const items =
                call === 0 ? [coveredRange] : textDeleted ? [] : [freshRange];
              return { items, load: vi.fn() };
            }),
            load: vi.fn(),
            text: "",
          },
          load: vi.fn(),
          changeTrackingMode: "off",
        },
        sync: vi.fn(async () => {}),
      };

      vi.stubGlobal("Word", {
        ChangeTrackingMode: { off: "off", trackAll: "trackAll", trackMine: "trackMine" },
        InsertLocation: { replace: "Replace" },
        run: vi.fn(async (cb: (ctx: typeof context) => unknown) => cb(context)),
      });

      const result = await new ApplySuggestionCommand(
        makeSuggestion(),
      ).execute();

      // No keepContent argument assertion — behavior proves it:
      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(freshRange.insertText).toHaveBeenCalled();
    });
  });

  describe("special character search", () => {
    it("Test C — searches for text starting with an em-dash", async () => {
      const originalText =
        "—¡Ah!, je, je, je, tendré más cuidado a partir de ahora.";
      const { context } = installWordContext();

      const command = new ApplySuggestionCommand(
        makeSuggestion({ originalText }),
      );
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(context.document.body.search).toHaveBeenCalledWith(originalText, {
        matchCase: true,
        matchWholeWord: false,
      });
    });

    it("Test D — searches with matchCase: true and matchWholeWord: false for em-dash and ellipsis", async () => {
      const originalText = "—¿Tú… en qué momento…?";
      const { context } = installWordContext();

      const command = new ApplySuggestionCommand(
        makeSuggestion({ originalText }),
      );
      await command.execute();

      expect(context.document.body.search).toHaveBeenCalledWith(originalText, {
        matchCase: true,
        matchWholeWord: false,
      });
      const callArgs = context.document.body.search.mock.calls[0] as [
        string,
        object,
      ];
      expect(callArgs[1]).toStrictEqual({
        matchCase: true,
        matchWholeWord: false,
      });
    });

    it("Test E — usa ignorePunct+ignoreSpace cuando exact match falla para em-dash (track-change)", async () => {
      const originalText =
        "—¡Ah!, je, je, je, tendré más cuidado a partir de ahora.";
      const { context, range } = installWordContext();

      context.document.body.search.mockReset();
      context.document.body.search
        .mockImplementationOnce(() => ({ items: [], load: vi.fn() })) // attempt 1: exact — fails
        .mockImplementationOnce(() => ({
          // attempt 1.5: ignorePunct+ignoreSpace — succeeds
          items: [
            {
              ...range,
              parentContentControlOrNullObject: {
                tag: "",
                isNullObject: true,
                load: vi.fn(),
                delete: vi.fn(),
              },
            },
          ],
          load: vi.fn(),
        }));

      const command = new ApplySuggestionCommand(
        makeSuggestion({ originalText }),
      );
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(context.document.body.search).toHaveBeenCalledTimes(2);
      expect(context.document.body.search).toHaveBeenNthCalledWith(
        1,
        originalText,
        {
          matchCase: true,
          matchWholeWord: false,
        },
      );
      expect(context.document.body.search).toHaveBeenNthCalledWith(
        2,
        originalText,
        {
          matchCase: true,
          matchWholeWord: false,
          ignorePunct: true,
          ignoreSpace: true,
        },
      );
      // Whitespace fallback body.load("text") must NOT have been called since 1.5 found it
      expect(context.document.body.load).not.toHaveBeenCalledWith("text");
    });

    it("Test F — usa ignorePunct+ignoreSpace en executeCommentOnly() también", async () => {
      const originalText = "—¡Ah, ¿qué es esto?!";
      const { context, range } = installWordContext();

      context.document.body.search.mockReset();
      context.document.body.search
        .mockImplementationOnce(() => ({ items: [], load: vi.fn() })) // attempt 1: exact — fails
        .mockImplementationOnce(() => ({
          // attempt 1.5: ignorePunct+ignoreSpace — succeeds
          items: [
            {
              ...range,
              parentContentControlOrNullObject: {
                tag: "",
                isNullObject: true,
                load: vi.fn(),
                delete: vi.fn(),
              },
            },
          ],
          load: vi.fn(),
        }));

      const command = new ApplySuggestionCommand(
        makeSuggestion({
          originalText,
          type: "comment-only",
          suggestedText: undefined,
        }),
      );
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(context.document.body.search).toHaveBeenCalledTimes(2);
      expect(context.document.body.search).toHaveBeenNthCalledWith(
        2,
        originalText,
        {
          matchCase: true,
          matchWholeWord: false,
          ignorePunct: true,
          ignoreSpace: true,
        },
      );
      // Whitespace fallback body.load("text") must NOT have been called since 1.5 found it
      expect(context.document.body.load).not.toHaveBeenCalledWith("text");
    });

    it("Test G — salta attempt 1 y va directo a attempt 1.5 si el texto supera 256 caracteres", async () => {
      const originalText = "a".repeat(257);
      const { context, range } = installWordContext();

      context.document.body.search.mockReset();
      context.document.body.search.mockImplementationOnce(() => ({
        // first actual call is attempt 1.5 (attempt 1 is skipped)
        items: [
          {
            ...range,
            parentContentControlOrNullObject: {
              tag: "",
              isNullObject: true,
              load: vi.fn(),
              delete: vi.fn(),
            },
          },
        ],
        load: vi.fn(),
      }));

      const command = new ApplySuggestionCommand(
        makeSuggestion({ originalText }),
      );
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      // Only one search call — the 1.5 attempt with ignorePunct+ignoreSpace
      expect(context.document.body.search).toHaveBeenCalledTimes(1);
      expect(context.document.body.search).toHaveBeenCalledWith(originalText, {
        matchCase: true,
        matchWholeWord: false,
        ignorePunct: true,
        ignoreSpace: true,
      });
    });
  });

  describe("comment-only suggestions", () => {
    it("inserts only a comment — no tracked change markup — when type is comment-only", async () => {
      const { context, cc } = installWordContext({
        initialTrackingMode: "trackAll",
      });

      const command = new ApplySuggestionCommand(
        makeSuggestion({
          id: "chunk1-5",
          type: "comment-only",
          suggestedText: undefined,
          originalText: "texto original",
          category: "Estilo",
          justification: "Mejora la claridad",
        }),
      );
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "chunk1-5" });

      const searchResult = context.document.body.search.mock.results[0]
        ?.value as SearchResults;
      const usedRange = searchResult.items[0];
      expect(usedRange.insertComment).toHaveBeenCalledWith(
        "[Estilo]\nMejora la claridad",
      );

      // No OOXML building — no insertText
      expect(usedRange.insertText).not.toHaveBeenCalled();

      // changeTrackingMode NOT touched
      expect(context.document.changeTrackingMode).toBe("trackAll");
      expect(context.document.load).not.toHaveBeenCalled();

      expect(cc.tag).toBe("stylistic:comment-only:chunk1-5");
    });

    it("returns a failed result when original text is not found for a comment-only suggestion", async () => {
      const { context } = installWordContext({ resultsItems: [] });

      const command = new ApplySuggestionCommand(
        makeSuggestion({ type: "comment-only", suggestedText: undefined }),
      );
      const result = await command.execute();

      expect(result).toEqual({
        success: false,
        commandId: "s1",
        error: "Texto original no encontrado",
      });
      expect(context.document.body.search).toHaveBeenCalledWith(
        "texto original",
        {
          matchCase: true,
          matchWholeWord: false,
        },
      );
    });

    it("does not mutate changeTrackingMode for comment-only suggestions", async () => {
      const { context } = installWordContext({
        initialTrackingMode: "trackAll",
      });

      const command = new ApplySuggestionCommand(
        makeSuggestion({ type: "comment-only", suggestedText: undefined }),
      );
      await command.execute();

      expect(context.document.load).not.toHaveBeenCalled();
      expect(context.document.changeTrackingMode).toBe("trackAll");
    });
  });
});
