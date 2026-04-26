import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "./WordAdapter";
import {
  makeCommentOnlyTag,
  installWordWithContext,
  makeOperationalWrapperTag,
  makeOperationalWrapperTitle,
  makeResolveSuggestionContext,
  makeSuggestion,
} from "./WordAdapterActionTestHelper";

describe("WordAdapter.acceptSuggestion", () => {
  let adapter: WordAdapter;

  beforeEach(() => {
    adapter = new WordAdapter();
    vi.spyOn(console, "log").mockImplementation(() => {});
    vi.spyOn(console, "warn").mockImplementation(() => {});
  });

  afterEach(() => {
    vi.restoreAllMocks();
    delete (globalThis as { Word?: unknown }).Word;
  });

  it("accepts a single-wrapper replace by resolving the wrapper range collection", async () => {
    const suggestion = makeSuggestion({
      id: "s-accept-single",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
    });
    const callOrder: string[] = [];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeOperationalWrapperTag("s-accept-single"),
      ccTitle: makeOperationalWrapperTitle({
        suggestionId: "s-accept-single",
        insertedTag: "stylistic:track-change:s-accept-single",
        deletedValue: "ni Shu",
        anchorValue:
          "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
      }),
      rangeTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(() => callOrder.push("accept-added")),
          reject: vi.fn(),
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(() => callOrder.push("accept-deleted")),
          reject: vi.fn(),
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder).toEqual(["accept-added", "accept-deleted"]);
  });

  it("returns ambiguous-location before mutating when duplicate valid wrappers exist", async () => {
    const suggestion = makeSuggestion({ id: "s-ambiguous-accept" });
    const unexpectedAccept = vi.fn();
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeOperationalWrapperTag("s-ambiguous-accept"),
      ccItems: [
        {
          tag: makeOperationalWrapperTag("s-ambiguous-accept"),
          title: makeOperationalWrapperTitle({
            suggestionId: "s-ambiguous-accept",
            insertedTag: "stylistic:track-change:s-ambiguous-accept",
          }),
          rangeTCItems: [
            {
              id: "tc-added-a",
              type: "Added",
              accept: unexpectedAccept,
              reject: vi.fn(),
            },
          ],
        },
        {
          tag: makeOperationalWrapperTag("s-ambiguous-accept"),
          title: makeOperationalWrapperTitle({
            suggestionId: "s-ambiguous-accept",
            insertedTag: "stylistic:track-change:s-ambiguous-accept",
          }),
          rangeTCItems: [
            {
              id: "tc-added-b",
              type: "Added",
              accept: unexpectedAccept,
              reject: vi.fn(),
            },
          ],
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("ambiguous-location");
    expect(result.trackedChangesAffected).toBe(0);
    expect(unexpectedAccept).not.toHaveBeenCalled();
  });

  it("degrades grouped wrappers before mutation instead of reconstructing host evidence", async () => {
    const suggestion = makeSuggestion({ id: "s-grouped-accept" });
    const unexpectedAccept = vi.fn();
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeOperationalWrapperTag("s-grouped-accept"),
      ccItems: [
        {
          tag: makeOperationalWrapperTag("s-grouped-accept"),
          title: makeOperationalWrapperTitle({
            suggestionId: "s-grouped-accept",
            insertedTag: "stylistic:track-change:s-grouped-accept",
            deletedValue: "texto original",
            anchorValue: "Contexto con texto original.",
            groupId: "group-a",
            groupIndex: 0,
            groupSize: 2,
          }),
          rangeTCItems: [
            {
              id: "tc-a",
              type: "Added",
              accept: unexpectedAccept,
              reject: vi.fn(),
            },
          ],
          rangeRelationWithNext: "AdjacentBefore",
        },
        {
          tag: makeOperationalWrapperTag("s-grouped-accept-1"),
          title: makeOperationalWrapperTitle({
            suggestionId: "s-grouped-accept-1",
            insertedTag: "stylistic:track-change:s-grouped-accept-1",
            deletedValue: "otro texto",
            anchorValue: "Otro contexto.",
            groupId: "group-a",
            groupIndex: 1,
            groupSize: 2,
          }),
          rangeTCItems: [
            {
              id: "tc-b",
              type: "Added",
              accept: unexpectedAccept,
              reject: vi.fn(),
            },
          ],
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("mixed-group");
    expect(result.trackedChangesAffected).toBe(0);
    expect(unexpectedAccept).not.toHaveBeenCalled();
  });

  it("accepts a comment-only suggestion by locating the canonical comment-only content control", async () => {
    const suggestion = makeSuggestion({
      id: "s-comment-only-accept",
      type: "comment-only",
      anchor: "frase observada",
      suggestedText: undefined,
      context: "Contexto con frase observada.",
    });
    const deleteComment = vi.fn();
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeCommentOnlyTag("s-comment-only-accept"),
      comments: [
        {
          authorName: "Usuario",
          content: "[Claridad]\nObservación",
          getRange: vi.fn(),
          delete: deleteComment,
        },
      ],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.commentDeleted).toBe(true);
    expect(context.document.contentControls.getByTag).toHaveBeenCalledWith(
      makeCommentOnlyTag("s-comment-only-accept"),
    );
    expect(context._cc.delete).toHaveBeenCalledWith(true);
    expect(deleteComment).toHaveBeenCalledOnce();
  });
});
