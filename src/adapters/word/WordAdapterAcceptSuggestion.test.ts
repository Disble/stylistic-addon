import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "./WordAdapter";
import {
  installRejectingWord,
  installWordWithContext,
  makeResolveSuggestionContext,
  makeSuggestion,
} from "./WordAdapterActionTestHelper";

describe("WordAdapter.acceptSuggestion", () => {
  let adapter: WordAdapter;
  let logSpy: ReturnType<typeof vi.spyOn>;
  let warnSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    adapter = new WordAdapter();
    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
    warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
  });

  afterEach(() => {
    logSpy.mockRestore();
    warnSpy.mockRestore();
    delete (globalThis as any).Word;
  });

  it("accepts both Stylistic tracked changes and deletes the colocated comment", async () => {
    const suggestion = makeSuggestion({
      id: "s-1",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });

    const tcAccept1 = vi.fn();
    const tcAccept2 = vi.fn();
    const commentDeleteSpy = vi.fn();

    const spanTCItems = [
      {
        author: "Stylistic",
        type: "Deleted",
        accept: tcAccept1,
        reject: vi.fn(),
      },
      {
        author: "Stylistic",
        type: "Added",
        accept: tcAccept2,
        reject: vi.fn(),
      },
    ];

    const commentRange = {
      compareLocationWith: vi.fn(() => ({ value: "Equal" })),
    };
    const comment = {
      authorName: "Stylistic",
      getRange: vi.fn(() => commentRange),
      delete: commentDeleteSpy,
    };

    const context = makeResolveSuggestionContext({
      ccFound: true,
      spanTCItems,
      comments: [comment],
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.commentDeleted).toBe(true);
    expect(tcAccept1).toHaveBeenCalledOnce();
    expect(tcAccept2).toHaveBeenCalledOnce();
    expect(commentDeleteSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("returns cc-not-found when the Content Control anchor is missing", async () => {
    const suggestion = makeSuggestion({
      anchor: "texto original",
      context: "Contexto con texto original.",
    });

    const context = makeResolveSuggestionContext({ ccFound: false });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("cc-not-found");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.commentDeleted).toBe(false);
  });

  it("returns already-resolved when the CC remains but no tracked changes are found", async () => {
    const suggestion = makeSuggestion({
      anchor: "texto original",
      context: "Contexto con texto original.",
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      spanTCItems: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("already-resolved");
    expect(result.trackedChangesAffected).toBe(0);
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("accepts tracked changes even when the associated comment is already gone", async () => {
    const suggestion = makeSuggestion({
      anchor: "texto original",
      context: "Contexto con texto original.",
    });
    const tcAcceptSpy = vi.fn();

    const spanTCItems = [
      {
        author: "Stylistic",
        type: "Deleted",
        accept: tcAcceptSpy,
        reject: vi.fn(),
      },
    ];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      spanTCItems,
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(1);
    expect(result.commentDeleted).toBe(false);
    expect(result.error).toBeUndefined();
    expect(tcAcceptSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("accepts overlapping body tracked changes when the CC-scoped collection misses one side", async () => {
    const suggestion = makeSuggestion({
      anchor: "quién",
      suggestedText: "quien",
      context: "Contexto con quién.",
    });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

    const ccScopedTrackedChanges = [
      {
        id: "tc-added",
        type: "Added",
        accept: addedAcceptSpy,
        reject: vi.fn(),
      },
    ];

    const bodyTrackedChanges = [
      ccScopedTrackedChanges[0],
      {
        id: "tc-deleted",
        type: "Deleted",
        accept: deletedAcceptSpy,
        reject: vi.fn(),
      },
    ];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      spanTCItems: ccScopedTrackedChanges,
      bodyTCItems: bodyTrackedChanges,
      bodyTCRelations: ["Equal", "OverlapsBefore"],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("accepts adjacent body tracked changes when the deleted side sits immediately before the CC", async () => {
    const suggestion = makeSuggestion({
      anchor: "cuál",
      suggestedText: "el cual",
      context: "Contexto con cuál.",
    });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

    const ccScopedTrackedChanges = [
      {
        id: "tc-added",
        type: "Added",
        accept: addedAcceptSpy,
        reject: vi.fn(),
      },
    ];

    const bodyTrackedChanges = [
      ccScopedTrackedChanges[0],
      {
        id: "tc-deleted",
        type: "Deleted",
        accept: deletedAcceptSpy,
        reject: vi.fn(),
      },
    ];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      spanTCItems: ccScopedTrackedChanges,
      bodyTCItems: bodyTrackedChanges,
      bodyTCRelations: ["Equal", "AdjacentBefore"],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("returns error without throwing when Word.run fails", async () => {
    const suggestion = makeSuggestion({
      anchor: "texto original",
      context: "Contexto con texto original.",
    });

    installRejectingWord(new Error("Document is read-only"));

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("error");
    expect(result.error).toContain("Document is read-only");
  });

  it("accepts comment-only suggestions by deleting the comment and the CC only", async () => {
    const suggestion = makeSuggestion({
      id: "s-co-1",
      type: "comment-only",
      suggestedText: undefined,
    });

    const commentDeleteSpy = vi.fn();
    const commentRange = {
      compareLocationWith: vi.fn(() => ({ value: "Equal" })),
    };
    const comment = {
      authorName: "Stylistic",
      getRange: vi.fn(() => commentRange),
      delete: commentDeleteSpy,
    };

    const context = makeResolveSuggestionContext({
      ccFound: true,
      spanTCItems: [],
      comments: [comment],
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.commentDeleted).toBe(true);
    expect(commentDeleteSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });
});
