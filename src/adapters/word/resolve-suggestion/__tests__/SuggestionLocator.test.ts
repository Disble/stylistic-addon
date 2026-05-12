import { describe, expect, it, vi } from "vitest";
import {
  installWordWithContext,
  makeCommentOnlyTag,
  makeOperationalWrapperTag,
  makeOperationalWrapperTitle,
  makeResolveSuggestionContext,
  makeSuggestion,
} from "../../__tests__/WordAdapterActionTestHelper";
import { SuggestionLocator } from "../SuggestionLocator";

describe("SuggestionLocator", () => {
  it("selects a unique strict operational wrapper without ranking legacy duplicates", async () => {
    const suggestion = makeSuggestion({ id: "s-1" });
    const locator = new SuggestionLocator(suggestion);
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccItems: [
        { tag: makeOperationalWrapperTag("s-1"), title: "legacy title" },
        {
          tag: makeOperationalWrapperTag("s-1"),
          title: makeOperationalWrapperTitle({
            suggestionId: suggestion.id,
            insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
            deletedValue: suggestion.anchor,
            anchorValue: suggestion.context,
          }),
        },
      ],
    });
    installWordWithContext(context);

    const located = await Word.run((wordContext) => locator.locateResolutionArtifacts(wordContext));

    expect(located.selectedCc).toBe(context._ccItems[1]);
    expect(located.locateStatus).toBe("confirmed-pending");
  });

  it("returns ambiguous-location when duplicate valid wrappers share one tag", async () => {
    const suggestion = makeSuggestion({ id: "s-dup" });
    const title = makeOperationalWrapperTitle({
      suggestionId: suggestion.id,
      insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
      deletedValue: suggestion.anchor,
      anchorValue: suggestion.context,
    });
    const locator = new SuggestionLocator(suggestion);
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccItems: [
        { tag: makeOperationalWrapperTag("s-dup"), title },
        { tag: makeOperationalWrapperTag("s-dup"), title },
      ],
    });
    installWordWithContext(context);

    const located = await Word.run((wordContext) => locator.locateResolutionArtifacts(wordContext));

    expect(located.selectedCc).toBeNull();
    expect(located.locateStatus).toBe("ambiguous-location");
  });

  it("returns ambiguous-location when only legacy wrappers are present", async () => {
    const suggestion = makeSuggestion({ id: "s-legacy" });
    const locator = new SuggestionLocator(suggestion);
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccItems: [{ tag: makeOperationalWrapperTag("s-legacy"), title: "legacy" }],
    });
    installWordWithContext(context);

    const located = await Word.run((wordContext) => locator.locateResolutionArtifacts(wordContext));

    expect(located.selectedCc).toBeNull();
    expect(located.locateStatus).toBe("ambiguous-location");
  });

  it("locates a unique comment-only content control using the canonical comment-only tag", async () => {
    const suggestion = makeSuggestion({
      id: "s-comment-only",
      type: "comment-only",
      suggestedText: undefined,
    });
    const locator = new SuggestionLocator(suggestion);
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccItems: [{ tag: makeCommentOnlyTag("s-comment-only"), title: suggestion.anchor }],
    });
    installWordWithContext(context);

    const located = await Word.run((wordContext) =>
      locator.locateCommentOnlyArtifacts(wordContext)
    );

    expect(located.selectedCc).toBe(context._ccItems[0]);
    expect(located.locateStatus).toBe("confirmed-pending");
  });

  it("returns cc-not-found when comment-only lookup misses the canonical tag", async () => {
    const suggestion = makeSuggestion({
      id: "s-comment-only-missing",
      type: "comment-only",
      suggestedText: undefined,
    });
    const locator = new SuggestionLocator(suggestion);
    const context = makeResolveSuggestionContext({ ccFound: false });
    installWordWithContext(context);

    const located = await Word.run((wordContext) =>
      locator.locateCommentOnlyArtifacts(wordContext)
    );

    expect(located.selectedCc).toBeNull();
    expect(located.locateStatus).toBe("cc-not-found");
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
    const comments = { items: [otherComment, matchingComment], load: vi.fn() };
    const context = {
      document: { body: { getComments: vi.fn(() => comments) } },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as Word.RequestContext;

    const found = await locator.findColocatedStylisticComment(context, cc);

    expect(found).toEqual({ comment: matchingComment, range: matchingRange });
  });
});
