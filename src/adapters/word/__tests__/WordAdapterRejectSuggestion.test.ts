import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "../WordAdapter";
import {
  makeCommentOnlyTag,
  installWordWithContext,
  makeOperationalWrapperTag,
  makeOperationalWrapperTitle,
  makeResolveSuggestionContext,
  makeSuggestion,
} from "../WordAdapterActionTestHelper";

describe("WordAdapter.rejectSuggestion", () => {
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

  it("rejects a single-wrapper replace by resolving the wrapper range collection", async () => {
    const suggestion = makeSuggestion({
      id: "s-reject-single",
      anchor: "eran meras suposiciones",
      suggestedText: "era una mera suposición",
      context: "Cualquier cosa que pudiera decir eran meras suposiciones.",
    });
    const callOrder: string[] = [];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeOperationalWrapperTag("s-reject-single"),
      ccTitle: makeOperationalWrapperTitle({
        suggestionId: "s-reject-single",
        insertedTag: "stylistic:track-change:s-reject-single",
        deletedValue: "eran meras suposiciones",
        anchorValue: "Cualquier cosa que pudiera decir eran meras suposiciones.",
      }),
      rangeTCItems: [
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-deleted")),
        },
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-added")),
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder).toEqual(["reject-deleted", "reject-added"]);
  });

  it("returns unobservable when the wrapper range exposes no actionable tracked changes", async () => {
    const suggestion = makeSuggestion({ id: "s-unobservable-reject" });
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeOperationalWrapperTag("s-unobservable-reject"),
      ccTitle: makeOperationalWrapperTitle({
        suggestionId: "s-unobservable-reject",
        insertedTag: "stylistic:track-change:s-unobservable-reject",
      }),
      rangeTCItems: [],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("unobservable");
    expect(result.trackedChangesAffected).toBe(0);
  });

  it("deletes only the located comment after a successful reject", async () => {
    const suggestion = makeSuggestion({ id: "s-comment-cleanup-reject" });
    const keptCommentDelete = vi.fn();
    const deletedCommentDelete = vi.fn();
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeOperationalWrapperTag("s-comment-cleanup-reject"),
      ccTitle: makeOperationalWrapperTitle({
        suggestionId: "s-comment-cleanup-reject",
        insertedTag: "stylistic:track-change:s-comment-cleanup-reject",
      }),
      rangeTCItems: [
        {
          id: "tc-only",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(),
        },
      ],
      comments: [
        {
          authorName: "User",
          content: "[Claridad]\nEliminar",
          getRange: vi.fn(),
          delete: deletedCommentDelete,
        },
        {
          authorName: "Otro",
          content: "No tocar",
          getRange: vi.fn(),
          delete: keptCommentDelete,
        },
      ],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(deletedCommentDelete).toHaveBeenCalledOnce();
    expect(keptCommentDelete).not.toHaveBeenCalled();
  });

  it("rejects a comment-only suggestion by locating the canonical comment-only content control", async () => {
    const suggestion = makeSuggestion({
      id: "s-comment-only-reject",
      type: "comment-only",
      anchor: "frase observada",
      suggestedText: undefined,
      context: "Contexto con frase observada.",
    });
    const deleteComment = vi.fn();
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeCommentOnlyTag("s-comment-only-reject"),
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

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.commentDeleted).toBe(true);
    expect(context.document.contentControls.getByTag).toHaveBeenCalledWith(
      makeCommentOnlyTag("s-comment-only-reject")
    );
    expect(context._cc.delete).toHaveBeenCalledWith(true);
    expect(deleteComment).toHaveBeenCalledOnce();
  });
});
