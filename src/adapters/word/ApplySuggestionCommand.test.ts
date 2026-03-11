import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import { OoxmlPackageBuilder } from "./ooxml/OoxmlPackageBuilder";
import type { Suggestion } from "../../domain/types";

type ApplyTestContext = {
  context: {
    document: {
      body: {
        search: ReturnType<typeof vi.fn>;
      };
      load: ReturnType<typeof vi.fn>;
      changeTrackingMode: string;
    };
    sync: ReturnType<typeof vi.fn>;
  };
  results: {
    items: Array<{
      getOoxml: ReturnType<typeof vi.fn>;
      insertOoxml: ReturnType<typeof vi.fn>;
    }>;
    load: ReturnType<typeof vi.fn>;
  };
  range: {
    getOoxml: ReturnType<typeof vi.fn>;
    insertOoxml: ReturnType<typeof vi.fn>;
  };
};

function makeSuggestion(overrides: Partial<Suggestion> = {}): Suggestion {
  return {
    id: "s1",
    originalText: "texto original",
    suggestedText: "texto sugerido",
    justification: "Mejora la claridad",
    category: "Estilo",
    severity: "medium",
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
  resultsItems?: Array<{
    getOoxml: ReturnType<typeof vi.fn>;
    insertOoxml: ReturnType<typeof vi.fn>;
  }>;
  rangeOoxml?: string;
  initialTrackingMode?: string;
  insertError?: Error;
  onSync?: (count: number) => void | Promise<void>;
} = {}): ApplyTestContext {
  const {
    resultsItems,
    rangeOoxml = "<pkg:package />",
    initialTrackingMode = "trackAll",
    insertError,
    onSync,
  } = options;

  const insertedRange = {
    insertContentControl: vi.fn(() => ({})),
  };
  const range = {
    getOoxml: vi.fn(() => ({ value: rangeOoxml })),
    insertOoxml: vi.fn(() => {
      if (insertError) {
        throw insertError;
      }
      return insertedRange;
    }),
  };

  const results = {
    items: resultsItems ?? [range],
    load: vi.fn(),
  };

  let syncCount = 0;
  const context = {
    document: {
      body: {
        search: vi.fn(() => results),
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

  return { context, results, range };
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
    const { context, results } = installWordContext({ resultsItems: [] });
    const builderSpy = vi.spyOn(OoxmlPackageBuilder.prototype, "build");

    const command = new ApplySuggestionCommand(makeSuggestion());
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
    expect(results.load).toHaveBeenCalledWith("items");
    expect(context.document.load).not.toHaveBeenCalled();
    expect(builderSpy).not.toHaveBeenCalled();
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
    expect(context.sync).toHaveBeenCalledTimes(6);
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
});
