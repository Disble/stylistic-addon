import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import { OoxmlPackageBuilder } from "./ooxml/OoxmlPackageBuilder";
import type { Suggestion } from "../../domain/types";

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
    getOoxml: ReturnType<typeof vi.fn>;
    insertOoxml: ReturnType<typeof vi.fn>;
    parentContentControlOrNullObject: ParentCC;
  };
  cc: {
    tag: string;
    appearance: string;
    cannotDelete: boolean;
  };
};

type SearchRange = {
  getOoxml: ReturnType<typeof vi.fn>;
  insertOoxml: ReturnType<typeof vi.fn>;
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

function installXmlMocks(options: {
  hasRunProperties?: boolean;
  serializedRunProperties?: string;
  throwOnParse?: boolean;
} = {}): void {
  const {
    hasRunProperties = true,
    serializedRunProperties = '<w:rPr><w:b/></w:rPr>',
    throwOnParse = false,
  } = options;

  vi.stubGlobal(
    "DOMParser",
    vi.fn().mockImplementation(() => ({
      parseFromString: vi.fn(() => {
        if (throwOnParse) {
          throw new Error("parse failed");
        }

        return {
          getElementsByTagNameNS: vi.fn(() => (hasRunProperties ? [{}] : [])),
        };
      }),
    }))
  );

  vi.stubGlobal(
    "XMLSerializer",
    vi.fn().mockImplementation(() => ({
      serializeToString: vi.fn(() => serializedRunProperties),
    }))
  );
}

function installWordContext(options: {
  resultsItems?: SearchRange[];
  searchResultsSequence?: SearchRange[][];
  documentText?: string;
  rangeOoxml?: string;
  initialTrackingMode?: string;
  insertError?: Error;
  onSync?: (count: number) => void | Promise<void>;
  parentCC?: Partial<ParentCC>;
} = {}): ApplyTestContext {
  const {
    resultsItems,
    searchResultsSequence,
    documentText = "",
    rangeOoxml = "<pkg:package />",
    initialTrackingMode = "trackAll",
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
  };
  const range: SearchRange = {
    getOoxml: vi.fn(() => ({ value: rangeOoxml })),
    insertOoxml: vi.fn(() => {
      if (insertError) {
        throw insertError;
      }
      return insertedRange;
    }),
    parentContentControlOrNullObject: parentCC,
  };

  const searchCalls: SearchResults[] = [];

  let syncCount = 0;
  const context = {
    document: {
      body: {
        search: vi.fn(() => {
          const nextItems = searchResultsSequence?.[searchCalls.length] ?? resultsItems ?? [range];
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
    },
    InsertLocation: {
      replace: "Replace",
    },
    run: vi.fn(async (callback: (ctx: typeof context) => unknown) => callback(context)),
  });

  return { context, range, cc };
}

describe("ApplySuggestionCommand", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
    vi.spyOn(console, "warn").mockImplementation(() => undefined);
    installXmlMocks();
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("returns a failed result when the original text is not found", async () => {
    const { context } = installWordContext({ resultsItems: [] });
    const builderSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "build");

    const command = new ApplySuggestionCommand(makeSuggestion());
    const result = await command.execute();

    const results = context.document.body.search.mock.results[0]?.value as SearchResults;

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Texto original no encontrado",
    });
    expect(context.document.body.search).toHaveBeenCalledWith("texto original", {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(results.load).toHaveBeenCalledWith("items");
    expect(context.document.load).not.toHaveBeenCalled();
    expect(builderSpy).not.toHaveBeenCalled();
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
      getOoxml: vi.fn(() => ({ value: "<pkg:package />" })),
      insertOoxml: range.insertOoxml,
      parentContentControlOrNullObject: { tag: "", isNullObject: true, load: vi.fn() },
    };
    context.document.body.search.mockReset();
    context.document.body.search
      .mockImplementationOnce(() => ({ items: [], load: vi.fn() }))
      .mockImplementationOnce(() => ({ items: [fallbackRange], load: vi.fn() }));

    const command = new ApplySuggestionCommand(makeSuggestion({ originalText }));
    const result = await command.execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(context.document.body.search).toHaveBeenNthCalledWith(1, originalText, {
      matchCase: true,
      matchWholeWord: false,
    });
    expect(context.document.body.load).toHaveBeenCalledWith("text");
    expect(context.document.body.search).toHaveBeenNthCalledWith(2, documentText, {
      matchCase: true,
      matchWholeWord: false,
    });
  });

  it("falls back when backend flattens paragraph breaks into spaces", async () => {
    const originalText =
      "los usuarios áuricos. —¿Eh?—Los humanos tampoco saben mucho sobre los humanos, ¿verdad? —dijo como si fuera una obviedad—.";
    const documentText =
      "los usuarios áuricos. \r—¿Eh?\r—Los humanos tampoco saben mucho sobre los humanos, ¿verdad? —dijo como si fuera una obviedad—.";

    const { context, range } = installWordContext({ documentText });
    const fallbackRange: SearchRange = {
      getOoxml: vi.fn(() => ({ value: "<pkg:package />" })),
      insertOoxml: range.insertOoxml,
      parentContentControlOrNullObject: { tag: "", isNullObject: true, load: vi.fn() },
    };
    context.document.body.search.mockReset();
    context.document.body.search
      .mockImplementationOnce(() => ({ items: [], load: vi.fn() }))
      .mockImplementationOnce(() => ({ items: [fallbackRange], load: vi.fn() }));

    const command = new ApplySuggestionCommand(makeSuggestion({ originalText }));
    const result = await command.execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(context.document.body.load).toHaveBeenCalledWith("text");
    expect(context.document.body.search).toHaveBeenNthCalledWith(2, documentText, {
      matchCase: true,
      matchWholeWord: false,
    });
  });

  it("builds OOXML for replace suggestions, inserts it, and restores tracking mode", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2025-04-05T06:07:08.999Z"));

    const { context, range } = installWordContext({
      rangeOoxml:
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:i/></w:rPr></w:document>',
    });
    const withRunPropertiesSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withRunProperties");
    const withChangeSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withChange");
    const withCommentSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withComment");

    const command = new ApplySuggestionCommand(makeSuggestion());
    const result = await command.execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(context.document.load).toHaveBeenCalledWith("changeTrackingMode");
    // QUIRK: in the node test runtime, run property extraction falls back to null
    // even when OOXML contains <w:rPr>, because the command swallows parser issues.
    expect(withRunPropertiesSpy).toHaveBeenCalledWith(null);
    expect(withChangeSpy).toHaveBeenCalledWith(
      "texto original",
      "texto sugerido",
      "replace",
      "Stylistic",
      "2025-04-05T06:07:08Z"
    );
    expect(withCommentSpy).toHaveBeenCalledWith(
      "Estilo",
      "Mejora la claridad",
      "Stylistic",
      "2025-04-05T06:07:08Z"
    );
    expect(range.insertOoxml).toHaveBeenCalledWith(
      expect.stringContaining('<w:delText xml:space="preserve">texto original</w:delText>'),
      "Replace"
    );
    expect(context.document.changeTrackingMode).toBe("trackAll");
    expect(context.sync).toHaveBeenCalledTimes(7);
  });

  it.each([
    {
      originalText: "texto original",
      suggestedText: "texto sugerido",
      expectedType: "replace",
      expectedSearchText: "texto original",
    },
    {
      originalText: "texto original",
      suggestedText: "",
      expectedType: "delete",
      expectedSearchText: "texto original",
    },
  ])(
    "classifies $expectedType suggestions through the public execute() path",
    async ({ originalText, suggestedText, expectedType, expectedSearchText }) => {
      const { context } = installWordContext();
      const withChangeSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withChange");

      const command = new ApplySuggestionCommand(
        makeSuggestion({ originalText, suggestedText })
      );

      await command.execute();

      expect(context.document.body.search).toHaveBeenCalledWith(expectedSearchText, {
        matchCase: true,
        matchWholeWord: false,
      });
      expect(withChangeSpy).toHaveBeenCalledWith(
        originalText,
        suggestedText,
        expectedType,
        "Stylistic",
        expect.any(String)
      );
    }
  );

  it("returns a failed result for insert-only suggestions without searching for empty text", async () => {
    const { context, range } = installWordContext();
    const withChangeSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withChange");
    const buildSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "build");

    const command = new ApplySuggestionCommand(
      makeSuggestion({ originalText: "", suggestedText: "texto sugerido" })
    );

    const result = await command.execute();

    expect(result).toEqual({
      success: false,
      commandId: "s1",
      error: "Insert-only suggestions require anchor text",
    });
    expect(context.document.body.search).not.toHaveBeenCalled();
    expect(context.document.load).not.toHaveBeenCalled();
    expect(range.getOoxml).not.toHaveBeenCalled();
    expect(range.insertOoxml).not.toHaveBeenCalled();
    expect(withChangeSpy).not.toHaveBeenCalled();
    expect(buildSpy).not.toHaveBeenCalled();
  });

  it("falls back to null run properties when OOXML parsing fails", async () => {
    installXmlMocks({ throwOnParse: true });
    installWordContext();
    const withRunPropertiesSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withRunProperties");

    const command = new ApplySuggestionCommand(makeSuggestion());
    const result = await command.execute();

    expect(result).toEqual({ success: true, commandId: "s1" });
    expect(withRunPropertiesSpy).toHaveBeenCalledWith(null);
  });

  it("returns a failed result when insertion throws and still restores tracking mode", async () => {
    const insertionError = new Error("insert failed");
    const { context } = installWordContext({ insertError: insertionError });

    const command = new ApplySuggestionCommand(makeSuggestion());

    await expect(command.execute()).resolves.toEqual({
      success: false,
      commandId: "s1",
      error: "insert failed",
    });
    expect(context.document.changeTrackingMode).toBe("trackAll");
  });

  it("assigns a namespaced CC tag with the suggestion type for track-change suggestions", async () => {
    const { cc } = installWordContext();

    const command = new ApplySuggestionCommand(makeSuggestion({ id: "chunk0-3" }));
    await command.execute();

    expect(cc.tag).toBe("stylistic:track-change:chunk0-3");
  });

  describe("already-covered detection (CC nesting prevention)", () => {
    function makeCleanRange(): SearchRange {
      return {
        getOoxml: vi.fn(() => ({ value: "<pkg:package />" })),
        insertOoxml: vi.fn(() => ({ insertContentControl: vi.fn(() => ({ tag: "", appearance: "", cannotDelete: true })) })),
        parentContentControlOrNullObject: { tag: "", isNullObject: true, load: vi.fn(), delete: vi.fn() },
      };
    }

    it("deletes the existing Stylistic CC, re-searches, and inserts OOXML for track-change", async () => {
      const coveredCC: ParentCC = { tag: "stylistic:track-change:chunk0-0", isNullObject: false, load: vi.fn(), delete: vi.fn() };
      const coveredRange: SearchRange = {
        getOoxml: vi.fn(() => ({ value: "<pkg:package />" })),
        insertOoxml: vi.fn(),
        parentContentControlOrNullObject: coveredCC,
      };
      const freshRange = makeCleanRange();
      const { context } = installWordContext({
        searchResultsSequence: [[coveredRange], [freshRange]],
      });

      const command = new ApplySuggestionCommand(makeSuggestion());
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(coveredCC.delete).toHaveBeenCalledWith(false);
      expect(context.document.body.search).toHaveBeenCalledTimes(2);
      expect(coveredRange.insertOoxml).not.toHaveBeenCalled();
      expect(freshRange.insertOoxml).toHaveBeenCalled();
    });

    it("deletes the existing legacy chunk CC, re-searches, and inserts OOXML", async () => {
      const coveredCC: ParentCC = { tag: "chunk0-0", isNullObject: false, load: vi.fn(), delete: vi.fn() };
      const coveredRange: SearchRange = {
        getOoxml: vi.fn(() => ({ value: "<pkg:package />" })),
        insertOoxml: vi.fn(),
        parentContentControlOrNullObject: coveredCC,
      };
      const freshRange = makeCleanRange();
      const { context } = installWordContext({
        searchResultsSequence: [[coveredRange], [freshRange]],
      });

      const command = new ApplySuggestionCommand(makeSuggestion());
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(coveredCC.delete).toHaveBeenCalledWith(false);
      expect(context.document.body.search).toHaveBeenCalledTimes(2);
      expect(coveredRange.insertOoxml).not.toHaveBeenCalled();
      expect(freshRange.insertOoxml).toHaveBeenCalled();
    });

    it("proceeds normally when range has no parent CC (isNullObject: true)", async () => {
      const { range } = installWordContext({
        parentCC: { isNullObject: true, tag: "" },
      });

      const command = new ApplySuggestionCommand(makeSuggestion());
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(range.parentContentControlOrNullObject.delete).not.toHaveBeenCalled();
      expect(range.insertOoxml).toHaveBeenCalled();
    });

    it("returns { success: false } when re-search after CC deletion finds nothing", async () => {
      const coveredCC: ParentCC = { tag: "stylistic:track-change:chunk0-0", isNullObject: false, load: vi.fn(), delete: vi.fn() };
      const coveredRange: SearchRange = {
        getOoxml: vi.fn(() => ({ value: "<pkg:package />" })),
        insertOoxml: vi.fn(),
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
    });

    it("deletes the existing Stylistic CC, re-searches, and inserts OOXML for comment-only", async () => {
      const coveredCC: ParentCC = { tag: "stylistic:comment-only:chunk1-5", isNullObject: false, load: vi.fn(), delete: vi.fn() };
      const coveredRange: SearchRange = {
        getOoxml: vi.fn(() => ({ value: "<pkg:package />" })),
        insertOoxml: vi.fn(),
        parentContentControlOrNullObject: coveredCC,
      };
      const freshRange = makeCleanRange();
      const { context } = installWordContext({
        searchResultsSequence: [[coveredRange], [freshRange]],
      });

      const command = new ApplySuggestionCommand(
        makeSuggestion({ type: "comment-only", suggestedText: undefined })
      );
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "s1" });
      expect(coveredCC.delete).toHaveBeenCalledWith(false);
      expect(context.document.body.search).toHaveBeenCalledTimes(2);
      expect(coveredRange.insertOoxml).not.toHaveBeenCalled();
      expect(freshRange.insertOoxml).toHaveBeenCalled();
    });
  });

  describe("special character search", () => {
    it("Test C — searches for text starting with an em-dash", async () => {
      const originalText = "—¡Ah!, je, je, je, tendré más cuidado a partir de ahora.";
      const { context } = installWordContext();

      const command = new ApplySuggestionCommand(makeSuggestion({ originalText }));
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

      const command = new ApplySuggestionCommand(makeSuggestion({ originalText }));
      await command.execute();

      expect(context.document.body.search).toHaveBeenCalledWith(originalText, {
        matchCase: true,
        matchWholeWord: false,
      });
      // Verify these exact options — no other matchXxx flags that could block special chars
      const callArgs = context.document.body.search.mock.calls[0] as [string, object];
      expect(callArgs[1]).toStrictEqual({ matchCase: true, matchWholeWord: false });
    });
  });

  describe("comment-only suggestions", () => {
    it("inserts only a comment — no tracked change markup — when type is comment-only", async () => {
      vi.useFakeTimers();
      vi.setSystemTime(new Date("2025-04-05T06:07:08.999Z"));

      const { cc } = installWordContext();
      const withChangeSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withChange");
      const withDeletionSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withDeletion");
      const withInsertionSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withInsertion");
      const withRunPropertiesSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withRunProperties");
      const withCommentSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withComment");

      const command = new ApplySuggestionCommand(
        makeSuggestion({
          id: "chunk1-5",
          type: "comment-only",
          suggestedText: undefined,
          originalText: "texto original",
          category: "Estilo",
          justification: "Mejora la claridad",
        })
      );
      const result = await command.execute();

      expect(result).toEqual({ success: true, commandId: "chunk1-5" });
      expect(withChangeSpy).not.toHaveBeenCalled();
      expect(withDeletionSpy).not.toHaveBeenCalled();
      expect(withInsertionSpy).not.toHaveBeenCalled();
      expect(withRunPropertiesSpy).not.toHaveBeenCalled();
      expect(withCommentSpy).toHaveBeenCalledWith(
        "Estilo",
        "Mejora la claridad",
        "Stylistic",
        "2025-04-05T06:07:08Z",
        "texto original"
      );
      expect(cc.tag).toBe("stylistic:comment-only:chunk1-5");
    });

    it("returns a failed result when original text is not found for a comment-only suggestion", async () => {
      const { context } = installWordContext({ resultsItems: [] });
      const buildSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "build");

      const command = new ApplySuggestionCommand(
        makeSuggestion({ type: "comment-only", suggestedText: undefined })
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
      expect(buildSpy).not.toHaveBeenCalled();
    });

    it("does not mutate changeTrackingMode for comment-only suggestions", async () => {
      const { context } = installWordContext({ initialTrackingMode: "trackAll" });

      const command = new ApplySuggestionCommand(
        makeSuggestion({ type: "comment-only", suggestedText: undefined })
      );
      await command.execute();

      expect(context.document.load).not.toHaveBeenCalled();
      expect(context.document.changeTrackingMode).toBe("trackAll");
    });

    it("passes originalText to withComment so the OOXML body preserves the matched text", async () => {
      vi.useFakeTimers();
      vi.setSystemTime(new Date("2025-04-05T06:07:08.999Z"));

      const { range } = installWordContext();
      const withCommentSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "withComment");

      const command = new ApplySuggestionCommand(
        makeSuggestion({
          id: "chunk2-7",
          type: "comment-only",
          suggestedText: undefined,
          originalText: "texto original",
          category: "Estilo",
          justification: "Mejora la claridad",
        })
      );
      await command.execute();

      expect(withCommentSpy).toHaveBeenCalledWith(
        "Estilo",
        "Mejora la claridad",
        "Stylistic",
        "2025-04-05T06:07:08Z",
        "texto original"
      );

      // The OOXML actually passed to insertOoxml must contain the originalText in the body
      const [insertedOoxml] = range.insertOoxml.mock.calls[0] as [string, string];
      const bodyStart = insertedOoxml.indexOf("<w:body>");
      const bodyEnd = insertedOoxml.indexOf("</w:body>");
      const body = insertedOoxml.slice(bodyStart, bodyEnd);
      expect(body).toContain("texto original");
    });
  });
});
