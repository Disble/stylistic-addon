import { describe, expect, it, vi } from "vitest";
import { SuggestionLocator } from "./SuggestionLocator";
import {
  installWordWithContext,
  makeCompoundV2Title,
  makeResolveSuggestionContext,
  makeSuggestion,
} from "../WordAdapterActionTestHelper";

describe("SuggestionLocator", () => {
  it("prefers the valid compound-v2 content control when duplicate tags exist", () => {
    const suggestion = makeSuggestion({ id: "s-1" });
    const locator = new SuggestionLocator(suggestion);
    const legacyCc = {
      tag: "stylistic:track-change:s-1",
      title: "legacy title",
    } as unknown as Word.ContentControl;
    const v2Cc = {
      tag: "stylistic:track-change:s-1",
      title: makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: suggestion.anchor,
        anchorValue: suggestion.context,
      }),
    } as unknown as Word.ContentControl;

    const selected = locator.selectResolutionContentControl([legacyCc, v2Cc]);

    expect(selected).toBe(v2Cc);
  });

  it("returns null when no valid compound-v2 content control exists", () => {
    const suggestion = makeSuggestion({ id: "s-1" });
    const locator = new SuggestionLocator(suggestion);
    const invalidCc = {
      tag: "stylistic:track-change:s-1",
      title: "legacy title",
    } as unknown as Word.ContentControl;

    const selected = locator.selectResolutionContentControl([invalidCc]);

    expect(selected).toBeNull();
  });

  it("keeps valid compound-v2 candidates first when ranking duplicates", () => {
    const suggestion = makeSuggestion({ id: "s-1" });
    const locator = new SuggestionLocator(suggestion);
    const legacyCc = {
      tag: "stylistic:track-change:s-1",
      title: "legacy title",
    } as unknown as Word.ContentControl;
    const v2Cc = {
      tag: "stylistic:track-change:s-1",
      title: makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: suggestion.anchor,
        anchorValue: suggestion.context,
      }),
    } as unknown as Word.ContentControl;

    const ranked = locator.rankResolutionContentControls([legacyCc, v2Cc]);

    expect(ranked[0]).toBe(v2Cc);
    expect(ranked[1]).toBe(legacyCc);
  });

  it("ranks the exact compound-v2 metadata match above structurally valid drifted duplicates", () => {
    const suggestion = makeSuggestion({
      id: "s-1",
      anchor: "fragmento actual",
      context: "Contexto con fragmento actual.",
    });
    const locator = new SuggestionLocator(suggestion);
    const driftedCc = {
      tag: "stylistic:track-change:s-1",
      title: makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: "anchor viejo",
        anchorValue: "Contexto viejo.",
      }),
    } as unknown as Word.ContentControl;
    const exactCc = {
      tag: "stylistic:track-change:s-1",
      title: makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: suggestion.anchor,
        anchorValue: suggestion.context,
      }),
    } as unknown as Word.ContentControl;

    const ranked = locator.rankResolutionContentControls([driftedCc, exactCc]);

    expect(ranked[0]).toBe(exactCc);
    expect(ranked[1]).toBe(driftedCc);
  });

  it("finds the colocated Stylistic comment that overlaps the selected content control", async () => {
    const suggestion = makeSuggestion({ id: "s-1" });
    const locator = new SuggestionLocator(suggestion);
    const ccRange = {} as Word.Range;
    const cc = {
      tag: "stylistic:track-change:s-1",
      getRange: vi.fn(() => ccRange),
    } as unknown as Word.ContentControl;
    const matchingRange = {
      compareLocationWith: vi.fn(() => ({ value: "Equal" })),
    } as unknown as Word.Range;
    const otherRange = {
      compareLocationWith: vi.fn(() => ({ value: "Before" })),
    } as unknown as Word.Range;
    const matchingComment = {
      authorName: "Usuario",
      content: "[Claridad]\nMas claro",
      getRange: vi.fn(() => matchingRange),
    } as unknown as Word.Comment;
    const otherComment = {
      authorName: "Usuario",
      content: "[Claridad]\nOtra cosa",
      getRange: vi.fn(() => otherRange),
    } as unknown as Word.Comment;
    const comments = {
      items: [otherComment, matchingComment],
      load: vi.fn(),
    };
    const context = {
      document: {
        body: {
          getComments: vi.fn(() => comments),
        },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as Word.RequestContext;

    const found = await locator.findColocatedStylisticComment(context, cc);

    expect(found).toEqual({ comment: matchingComment, range: matchingRange });
  });

  it("logs structured candidate diagnostics when duplicate tagged content controls exist", async () => {
    const suggestion = makeSuggestion({
      id: "s-dup-log",
      anchor: "desde allí",
      context: "Contexto con desde allí.",
    });
    const locator = new SuggestionLocator(suggestion);
    const consoleLogSpy = vi.spyOn(console, "log").mockImplementation(() => {});
    const consoleWarnSpy = vi
      .spyOn(console, "warn")
      .mockImplementation(() => {});
    const duplicateTitle = makeCompoundV2Title({
      suggestionId: suggestion.id,
      insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
      deletedValue: suggestion.anchor,
      anchorValue: suggestion.context,
    });
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-dup-log",
      ccItems: [
        {
          tag: "stylistic:track-change:s-dup-log",
          title: duplicateTitle,
        },
        {
          tag: "stylistic:track-change:s-dup-log",
          title: duplicateTitle,
        },
      ],
    });
    installWordWithContext(context);

    const located = await Word.run((wordContext) =>
      locator.locateResolutionArtifacts(wordContext),
    );

    expect(located.selectedCc).toBe(context._ccItems[0]);
    expect(consoleLogSpy).toHaveBeenCalledWith(
      `🧾 [SuggestionLocator] candidate diagnostics for suggestionId="${suggestion.id}"`,
      expect.arrayContaining([
        expect.objectContaining({
          candidateIndex: 0,
          wasSelected: true,
          selectionReason: "valid-compound-v2",
          titleKind: "compound-v2",
          validCompoundV2: true,
          score: 3,
          identitySuggestionId: suggestion.id,
          deletedSideValue: suggestion.anchor,
          anchorValue: suggestion.context,
        }),
        expect.objectContaining({
          candidateIndex: 1,
          wasSelected: false,
          selectionReason: "not-selected",
          titleKind: "compound-v2",
          validCompoundV2: true,
          score: 3,
          identitySuggestionId: suggestion.id,
        }),
      ]),
    );
    expect(consoleWarnSpy).toHaveBeenCalledWith(
      `⚠️ [SuggestionLocator] indistinguishable duplicate candidates for suggestionId="${suggestion.id}"`,
      expect.arrayContaining([
        expect.objectContaining({ candidateIndex: 0 }),
        expect.objectContaining({ candidateIndex: 1 }),
      ]),
    );

    consoleLogSpy.mockRestore();
    consoleWarnSpy.mockRestore();
  });
});
