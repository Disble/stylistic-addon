import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "./WordAdapter";
import {
  installWordWithContext,
  makeResolveSuggestionContext,
  makeSuggestion,
} from "./WordAdapterActionTestHelper";

describe("WordAdapter.rejectSuggestion", () => {
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

  it("rejects both Stylistic tracked changes and deletes the colocated comment", async () => {
    const suggestion = makeSuggestion({
      id: "s-1",
      anchor: "texto original",
      context: "Contexto con texto original.",
    });

    const tcReject1 = vi.fn();
    const tcReject2 = vi.fn();
    const tcAccept1 = vi.fn();
    const tcAccept2 = vi.fn();
    const commentDeleteSpy = vi.fn();

    const spanTCItems = [
      {
        author: "Stylistic",
        type: "Deleted",
        accept: tcAccept1,
        reject: tcReject1,
      },
      {
        author: "Stylistic",
        type: "Added",
        accept: tcAccept2,
        reject: tcReject2,
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

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(tcReject1).toHaveBeenCalledOnce();
    expect(tcReject2).toHaveBeenCalledOnce();
    expect(tcAccept1).not.toHaveBeenCalled();
    expect(tcAccept2).not.toHaveBeenCalled();
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

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("cc-not-found");
  });

  it("rejects comment-only suggestions by deleting the comment and the CC only", async () => {
    const suggestion = makeSuggestion({
      id: "s-co-2",
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

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.commentDeleted).toBe(true);
    expect(commentDeleteSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("rejects overlapping body tracked changes when the CC-scoped collection misses one side", async () => {
    const suggestion = makeSuggestion({
      anchor: "quién",
      suggestedText: "quien",
      context: "Contexto con quién.",
    });
    const addedRejectSpy = vi.fn();
    const deletedRejectSpy = vi.fn();

    const ccScopedTrackedChanges = [
      {
        id: "tc-added",
        type: "Added",
        accept: vi.fn(),
        reject: addedRejectSpy,
      },
    ];

    const bodyTrackedChanges = [
      ccScopedTrackedChanges[0],
      {
        id: "tc-deleted",
        type: "Deleted",
        accept: vi.fn(),
        reject: deletedRejectSpy,
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

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedRejectSpy).toHaveBeenCalledOnce();
    expect(deletedRejectSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("rejects adjacent body tracked changes when the deleted side sits immediately before the CC", async () => {
    const suggestion = makeSuggestion({
      anchor: "cuál",
      suggestedText: "el cual",
      context: "Contexto con cuál.",
    });
    const addedRejectSpy = vi.fn();
    const deletedRejectSpy = vi.fn();

    const ccScopedTrackedChanges = [
      {
        id: "tc-added",
        type: "Added",
        accept: vi.fn(),
        reject: addedRejectSpy,
      },
    ];

    const bodyTrackedChanges = [
      ccScopedTrackedChanges[0],
      {
        id: "tc-deleted",
        type: "Deleted",
        accept: vi.fn(),
        reject: deletedRejectSpy,
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

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedRejectSpy).toHaveBeenCalledOnce();
    expect(deletedRejectSpy).toHaveBeenCalledOnce();
  });

  it("returns rejected even when rejecting the tracked changes invalidates the CC before final cleanup", async () => {
    const suggestion = makeSuggestion({
      anchor: "con la Jing",
      suggestedText: "con Jing",
      context: "Contexto con la Jing.",
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      spanTCItems: [],
      comments: [],
    });

    const rejectAddedSpy = vi.fn(() => {
      context._cc.delete.mockImplementation(() => {
        throw new Error("GeneralException");
      });
    });

    const rejectDeletedSpy = vi.fn();

    const ccScopedTrackedChanges = [
      {
        id: "tc-added",
        type: "Added",
        accept: vi.fn(),
        reject: rejectAddedSpy,
      },
      {
        id: "tc-deleted",
        type: "Deleted",
        accept: vi.fn(),
        reject: rejectDeletedSpy,
      },
    ];

    context._cc.getTrackedChanges.mockReturnValue({
      items: ccScopedTrackedChanges,
      load: vi.fn(),
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(rejectAddedSpy).toHaveBeenCalledOnce();
    expect(rejectDeletedSpy).toHaveBeenCalledOnce();
  });
});
