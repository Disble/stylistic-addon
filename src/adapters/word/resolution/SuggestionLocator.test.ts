import { describe, expect, it, vi } from "vitest";
import { SuggestionLocator } from "./SuggestionLocator";
import {
  makeCompoundV2Title,
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
});
