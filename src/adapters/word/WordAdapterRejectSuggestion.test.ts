import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "./WordAdapter";
import {
  installWordWithContext,
  makeCompoundV2Title,
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
      authorName: "Usuario de prueba",
      content: "[Claridad]\nMas claro",
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
      authorName: "Usuario de prueba",
      content: "[Claridad]\nMas claro",
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

  it("returns unobservable when reject cannot observe any tracked changes for the CC", async () => {
    const suggestion = makeSuggestion({
      anchor: "texto original",
      context: "Contexto con texto original.",
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      spanTCItems: [],
      bodyTCItems: [],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("unobservable");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.error).toContain("Word no expuso suficientes tracked changes");
    expect(context._cc.delete).not.toHaveBeenCalled();
  });

  it("rejects compound-v2 replace suggestions when rich metadata and tracked changes are present", async () => {
    const suggestion = makeSuggestion({ id: "s-1" });
    const addedRejectSpy = vi.fn();
    const deletedRejectSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTitle: makeCompoundV2Title(),
      spanTCItems: [
        { id: "tc-added", type: "Added", accept: vi.fn(), reject: addedRejectSpy },
        { id: "tc-deleted", type: "Deleted", accept: vi.fn(), reject: deletedRejectSpy },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedRejectSpy).toHaveBeenCalledOnce();
    expect(deletedRejectSpy).toHaveBeenCalledOnce();
  });

  it("returns identity-lost for corrupt compound-v2 metadata during reject", async () => {
    const suggestion = makeSuggestion({ id: "s-1" });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTitle: makeCompoundV2Title({ overrides: { insertedSideRef: { kind: "anchor", role: "inserted-side", value: "bad" } } }),
      spanTCItems: [],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("identity-lost");
    expect(result.trackedChangesAffected).toBe(0);
    expect(context._cc.delete).not.toHaveBeenCalled();
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
      ccTitle: makeCompoundV2Title({
        deletedValue: "quién",
        anchorValue: "Contexto con quién.",
      }),
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
      ccTitle: makeCompoundV2Title({
        deletedValue: "cuál",
        anchorValue: "Contexto con cuál.",
      }),
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

  it("rejects replace suggestions when only the CC range exposes tracked changes in real-host style semantics", async () => {
    const suggestion = makeSuggestion({
      anchor: "quién",
      suggestedText: "quien",
      context: "Contexto con quién.",
    });
    const addedRejectSpy = vi.fn();
    const deletedRejectSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTitle: makeCompoundV2Title({
        deletedValue: "quién",
        anchorValue: "Contexto con quién.",
      }),
      spanTCItems: [],
      rangeTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(),
          reject: addedRejectSpy,
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(),
          reject: deletedRejectSpy,
        },
      ],
      bodyTCItems: [],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedRejectSpy).toHaveBeenCalledOnce();
    expect(deletedRejectSpy).toHaveBeenCalledOnce();
  });

  it("rejects replace suggestions when only the operational anchor range exposes tracked changes", async () => {
    const suggestion = makeSuggestion({
      id: "s-anchor-only",
      anchor: "parecía ser",
      suggestedText: "parece que es",
      context: "Hasta donde sabía, parecía ser la única al tanto.",
    });
    const addedRejectSpy = vi.fn();
    const deletedRejectSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-anchor-only",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-anchor-only",
        insertedTag: "stylistic:track-change:s-anchor-only",
        deletedValue: "parecía ser",
        anchorValue: "Hasta donde sabía, parecía ser la única al tanto.",
      }),
      spanTCItems: [],
      rangeTCItems: [],
      bodyTCItems: [],
      operationalAnchorText: "Hasta donde sabía, parecía ser la única al tanto.",
      operationalAnchorRangeTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(),
          reject: addedRejectSpy,
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(),
          reject: deletedRejectSpy,
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedRejectSpy).toHaveBeenCalledOnce();
    expect(deletedRejectSpy).toHaveBeenCalledOnce();
  });

  it("rejects replace suggestions when only the colocated comment range exposes tracked changes", async () => {
    const suggestion = makeSuggestion({
      id: "s-comment-range",
      anchor: "parecía ser",
      suggestedText: "parece que es",
      context: "Hasta donde sabía, parecía ser la única al tanto.",
      category: "Gramática",
      justification: "Ajuste verbal",
    });
    const addedRejectSpy = vi.fn();
    const deletedRejectSpy = vi.fn();

    const commentRange = {
      compareLocationWith: vi.fn(() => ({ value: "Equal" })),
    };

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-comment-range",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-comment-range",
        insertedTag: "stylistic:track-change:s-comment-range",
        deletedValue: "parecía ser",
        anchorValue: "Hasta donde sabía, parecía ser la única al tanto.",
      }),
      spanTCItems: [],
      rangeTCItems: [],
      bodyTCItems: [],
      operationalAnchorText: undefined,
      comments: [
        {
          authorName: "Stylistic",
          content: "[Gramática]\nAjuste verbal",
          getRange: vi.fn(() => commentRange),
          delete: vi.fn(),
        },
      ],
      commentRangeTCItems: [
        [
          {
            id: "tc-added",
            type: "Added",
            accept: vi.fn(),
            reject: addedRejectSpy,
          },
          {
            id: "tc-deleted",
            type: "Deleted",
            accept: vi.fn(),
            reject: deletedRejectSpy,
          },
        ],
      ],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedRejectSpy).toHaveBeenCalledOnce();
    expect(deletedRejectSpy).toHaveBeenCalledOnce();
  });

  it("prefers the compound-v2 CC when multiple CCs share the same tag", async () => {
    const suggestion = makeSuggestion({ id: "chunk0-0" });
    const addedRejectSpy = vi.fn();
    const deletedRejectSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-0",
      ccItems: [
        {
          tag: "stylistic:track-change:chunk0-0",
          title: "texto legacy sin metadata",
          spanTCItems: [],
          rangeTCItems: [],
        },
        {
          tag: "stylistic:track-change:chunk0-0",
          title: makeCompoundV2Title({
            suggestionId: "chunk0-0",
            insertedTag: "stylistic:track-change:chunk0-0",
          }),
          spanTCItems: [
            {
              id: "tc-added",
              type: "Added",
              accept: vi.fn(),
              reject: addedRejectSpy,
            },
            {
              id: "tc-deleted",
              type: "Deleted",
              accept: vi.fn(),
              reject: deletedRejectSpy,
            },
          ],
          rangeTCItems: [],
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedRejectSpy).toHaveBeenCalledOnce();
    expect(deletedRejectSpy).toHaveBeenCalledOnce();
  });

  it("tries later CC candidates when the first v2 candidate remains unobservable", async () => {
    const suggestion = makeSuggestion({ id: "chunk0-0" });
    const addedRejectSpy = vi.fn();
    const deletedRejectSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-0",
      ccItems: [
        {
          tag: "stylistic:track-change:chunk0-0",
          title: makeCompoundV2Title({
            suggestionId: "chunk0-0",
            insertedTag: "stylistic:track-change:chunk0-0",
          }),
          spanTCItems: [],
          rangeTCItems: [],
        },
        {
          tag: "stylistic:track-change:chunk0-0",
          title: makeCompoundV2Title({
            suggestionId: "chunk0-0",
            insertedTag: "stylistic:track-change:chunk0-0",
          }),
          spanTCItems: [
            {
              id: "tc-added",
              type: "Added",
              accept: vi.fn(),
              reject: addedRejectSpy,
            },
            {
              id: "tc-deleted",
              type: "Deleted",
              accept: vi.fn(),
              reject: deletedRejectSpy,
            },
          ],
          rangeTCItems: [],
        },
      ],
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
      ccTitle: makeCompoundV2Title({
        deletedValue: "con la Jing",
        anchorValue: "Contexto con la Jing.",
      }),
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

  it("returns rejected when reject succeeds but a later document-state read throws GeneralException", async () => {
    const suggestion = makeSuggestion({ id: "chunk0-7" });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-7",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-7",
        insertedTag: "stylistic:track-change:chunk0-7",
      }),
      spanTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(),
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(),
        },
      ],
      comments: [],
    });

    let rejectExecuted = false;
    context._cc.getTrackedChanges.mockReturnValue({
      items: [
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(() => {
            rejectExecuted = true;
          }),
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(),
        },
      ],
      load: vi.fn(),
    });

    context.document.contentControls.load.mockImplementation(() => {
      if (rejectExecuted) {
        throw new Error("GeneralException");
      }
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(rejectExecuted).toBe(true);
    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
  });

  it("signals the disable CTA when rejecting the final pending artifact reaches zero", async () => {
    const suggestion = makeSuggestion({ id: "s-final" });
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-final",
      spanTCItems: [
        {
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(() => {
            context.document.contentControls.items = [];
          }),
        },
      ],
      comments: [],
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.pendingAfter.pendingStylisticArtifacts).toBe(0);
    expect(result.documentState).toBe("ready-to-disable-track-changes");
  });

  it("returns rejected (not error) when rejecting tracked changes causes comment.delete() sync to throw GeneralException", async () => {
    // Reproduces the real-host bug: tc.reject() invalidates the Word context, so
    // the context.sync() inside deleteLocatedStylisticComment throws GeneralException
    // AFTER the tracked changes were already successfully rejected in Word.
    // Expected: status="rejected" so the taskpane UI updates correctly.
    // Actual (before fix): status="error" because the exception escaped to the outer catch.
    const suggestion = makeSuggestion({
      id: "chunk0-3",
      anchor: "sería a quien lo desconocía",
      suggestedText: "sería quien lo desconocía",
      context: "sería a quien lo desconocía. Su vida cambiaría.",
    });

    const rejectSpy1 = vi.fn();
    const rejectSpy2 = vi.fn();
    let rejectExecuted = false;

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-3",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-3",
        insertedTag: "stylistic:track-change:chunk0-3",
        deletedValue: "sería a quien lo desconocía",
        anchorValue: "sería a quien lo desconocía. Su vida cambiaría.",
      }),
      spanTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(() => {
            rejectExecuted = true;
            rejectSpy1();
          }),
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(),
          reject: rejectSpy2,
        },
      ],
      comments: [
        {
          authorName: "Stylistic",
          content: "[gramática]\nsería quien lo desconocía",
          getRange: vi.fn(() => ({ compareLocationWith: vi.fn(() => ({ value: "Equal" })) })),
          delete: vi.fn(),
        },
      ],
    });

    // Simulate: after tc.reject() executes, the context.sync() throws GeneralException
    // (this is what happens in real Word when reject invalidates the comment/CC proxy)
    context.sync.mockImplementation(() => {
      if (rejectExecuted) {
        return Promise.reject(new Error("GeneralException"));
      }
      return Promise.resolve();
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(rejectSpy1).toHaveBeenCalledOnce();
    expect(rejectSpy2).toHaveBeenCalledOnce();
  });
});
