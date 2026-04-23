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

  it("rejects replace tracked changes by restoring deletion before removing insertion", async () => {
    const suggestion = makeSuggestion({
      id: "s-ordered-reject",
      anchor: "desde allí",
      suggestedText: "desde entonces",
      context:
        "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
    });
    const callOrder: string[] = [];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-ordered-reject",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-ordered-reject",
        insertedTag: "stylistic:track-change:s-ordered-reject",
        deletedValue: "desde allí",
        anchorValue:
          "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
      }),
      spanTCItems: [
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(() => {
            callOrder.push("reject-deleted");
          }),
        },
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(() => {
            callOrder.push("reject-added");
          }),
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

  it("syncs after each replace reject step before continuing with the next semantic side", async () => {
    const suggestion = makeSuggestion({
      id: "s-sync-reject",
      anchor: "desde allí",
      suggestedText: "desde entonces",
      context:
        "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
    });
    const callOrder: string[] = [];
    let queuedStep = 0;
    let lastSyncedStep = 0;

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-sync-reject",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-sync-reject",
        insertedTag: "stylistic:track-change:s-sync-reject",
        deletedValue: "desde allí",
        anchorValue:
          "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
      }),
      spanTCItems: [
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(() => {
            queuedStep = 1;
            callOrder.push("reject-deleted");
          }),
        },
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(() => {
            queuedStep = 2;
            callOrder.push("reject-added");
          }),
        },
      ],
      comments: [],
    });

    context.sync.mockImplementation(async () => {
      if (queuedStep > lastSyncedStep) {
        callOrder.push("sync");
        lastSyncedStep = queuedStep;
      }
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder.slice(0, 4)).toEqual([
      "reject-deleted",
      "sync",
      "reject-added",
      "sync",
    ]);
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

  it("returns unobservable when a replace suggestion exposes only the inserted side during reject", async () => {
    const suggestion = makeSuggestion({
      id: "s-half-visible",
      anchor: "desde allí",
      suggestedText: "desde entonces",
      context:
        "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
    });
    const addedRejectSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-half-visible",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-half-visible",
        insertedTag: "stylistic:track-change:s-half-visible",
        deletedValue: "desde allí",
        anchorValue:
          "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
      }),
      spanTCItems: [
        {
          id: "tc-added-only",
          type: "Added",
          accept: vi.fn(),
          reject: addedRejectSpy,
        },
      ],
      rangeTCItems: [],
      bodyTCItems: [],
      comments: [],
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("unobservable");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.commentDeleted).toBe(false);
    expect(result.error).toContain("Word no expuso suficientes tracked changes");
    expect(addedRejectSpy).not.toHaveBeenCalled();
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

  it("returns identity-lost for structurally corrupt compound-v2 metadata during reject", async () => {
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

  it("still resolves reject when the only compound-v2 artifact has host-drifted deleted and anchor text", async () => {
    const suggestion = makeSuggestion({
      id: "s-1",
      anchor: "texto actual",
      context: "Contexto con texto actual.",
    });
    const addedRejectSpy = vi.fn();
    const deletedRejectSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTitle: makeCompoundV2Title({
        suggestionId: suggestion.id,
        insertedTag: `stylistic:${suggestion.type}:${suggestion.id}`,
        deletedValue: "texto viejo preservado por Word",
        anchorValue: "Contexto viejo rehidratado por Word.",
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
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedRejectSpy).toHaveBeenCalledOnce();
    expect(deletedRejectSpy).toHaveBeenCalledOnce();
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

  it("does not over-collect replace tracked changes during reject when stale sources expose extras", async () => {
    const suggestion = makeSuggestion({
      id: "s-no-overcollect",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });
    const callOrder: string[] = [];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-no-overcollect",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-no-overcollect",
        insertedTag: "stylistic:track-change:s-no-overcollect",
        deletedValue: "ni Shu",
        anchorValue: "No sabían si venía de ni Shu o de otro sitio.",
      }),
      spanTCItems: [
        {
          id: "tc-added-main",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-added-main")),
        },
      ],
      rangeTCItems: [],
      bodyTCItems: [
        {
          id: "tc-added-main",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-added-duplicate")),
        },
        {
          id: "tc-deleted-stale-body",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-deleted-stale-body")),
        },
      ],
      bodyTCRelations: ["Equal", "AdjacentBefore"],
      deletedSideText: "ni Shu",
      deletedSideRangeTCItems: [
        {
          id: "tc-deleted-main",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-deleted-main")),
        },
      ],
      operationalAnchorText: "No sabían si venía de ni Shu o de otro sitio.",
      operationalAnchorRangeTCItems: [
        {
          id: "tc-added-stale-anchor",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-added-stale-anchor")),
        },
        {
          id: "tc-deleted-stale-anchor",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-deleted-stale-anchor")),
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    // Deleted-first semantic order: cc-internal Deleted is deprioritized so the
    // first step picks the deletedSide locator (tc-deleted-main). The inter-step
    // re-observation for the remaining Added side then picks tc-added-main from
    // the cc/ccRange scope (its body duplicate is id-deduped out of bodyRelated).
    expect(callOrder).toEqual([
      "reject-deleted-main",
      "reject-added-main",
    ]);
  });

  it("normalizes duplicate deleted-side observations into a single semantic replace pair during reject", async () => {
    const suggestion = makeSuggestion({
      id: "s-semantic-dedupe",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
    });
    const callOrder: string[] = [];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-semantic-dedupe",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-semantic-dedupe",
        insertedTag: "stylistic:track-change:s-semantic-dedupe",
        deletedValue: "ni Shu",
        anchorValue:
          "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
      }),
      spanTCItems: [
        {
          id: "tc-deleted-cc",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-deleted-cc")),
        },
      ],
      rangeTCItems: [
        {
          id: "tc-added-range",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-added-range")),
        },
      ],
      bodyTCItems: [
        {
          id: "tc-deleted-body",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(() => callOrder.push("reject-deleted-body")),
        },
      ],
      bodyTCRelations: ["AdjacentBefore"],
      comments: [
        {
          authorName: "Stylistic",
          content: "[gramática]\nAjuste",
          getRange: vi.fn(() => ({ compareLocationWith: vi.fn(() => ({ value: "Equal" })) })),
          delete: vi.fn(),
        },
      ],
      commentRangeTCItems: [
        [
          {
            id: "tc-deleted-comment",
            type: "Deleted",
            accept: vi.fn(),
            reject: vi.fn(() => callOrder.push("reject-deleted-comment")),
          },
        ],
      ],
    });
    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    // Deleted-first semantic order: bodyRelated Deleted is rejected first, then
    // ccRange Added. The cc-internal Deleted is skipped (cc.getTrackedChanges is
    // deprioritized).
    expect(callOrder).toEqual(["reject-deleted-body", "reject-added-range"]);
  });

  it("rejects replace suggestions when the deleted-side locator exposes the missing deleted tracked change", async () => {
    const suggestion = makeSuggestion({
      id: "s-pre-replace-context",
      anchor: "parecía ser",
      suggestedText: "parece que es",
      context: "Hasta donde sabía, parecía ser la única al tanto.",
    });
    const addedRejectSpy = vi.fn();
    const deletedRejectSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-pre-replace-context",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-pre-replace-context",
        insertedTag: "stylistic:track-change:s-pre-replace-context",
        deletedValue: "parecía ser",
        anchorValue: "Hasta donde sabía, parecía ser la única al tanto.",
      }),
      spanTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(),
          reject: addedRejectSpy,
        },
      ],
      rangeTCItems: [],
      bodyTCItems: [],
      deletedSideText: "parecía ser",
      deletedSideRangeTCItems: [
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(),
          reject: deletedRejectSpy,
        },
      ],
      operationalAnchorText: undefined,
      operationalAnchorRangeTCItems: [],
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

  it("returns error when rejecting the tracked changes invalidates the CC before final cleanup", async () => {
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

    const rejectDeletedSpy = vi.fn(() => {
      resolutionPhase = 1;
    });

    let resolutionPhase = 0;
    const rejectAddedSpy = vi.fn(() => {
      resolutionPhase = 2;
      context._cc.delete.mockImplementation(() => {
        throw new Error("GeneralException");
      });
    });

    context._cc.getTrackedChanges.mockImplementation(() => {
      if (resolutionPhase === 0) {
        return {
          items: [
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
          ],
          load: vi.fn(),
        };
      }

      if (resolutionPhase === 1) {
        return {
          items: [
            {
              id: "tc-added",
              type: "Added",
              accept: vi.fn(),
              reject: rejectAddedSpy,
            },
          ],
          load: vi.fn(),
        };
      }

      return {
        items: [],
        load: vi.fn(),
      };
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("error");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.error).toBe("GeneralException");
    expect(rejectAddedSpy).toHaveBeenCalledOnce();
    expect(rejectDeletedSpy).toHaveBeenCalledOnce();
  });

  it("returns error when reject succeeds but a later document-state read throws GeneralException", async () => {
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
    let resolutionPhase = 0;
    const rejectDeletedSpy = vi.fn(() => {
      resolutionPhase = 1;
    });
    const rejectAddedSpy = vi.fn(() => {
      rejectExecuted = true;
      resolutionPhase = 2;
    });
    context._cc.getTrackedChanges.mockImplementation(() => {
      if (resolutionPhase === 0) {
        return {
          items: [
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
          ],
          load: vi.fn(),
        };
      }

      if (resolutionPhase === 1) {
        return {
          items: [
            {
              id: "tc-added",
              type: "Added",
              accept: vi.fn(),
              reject: rejectAddedSpy,
            },
          ],
          load: vi.fn(),
        };
      }

      return {
        items: [],
        load: vi.fn(),
      };
    });

    context.document.contentControls.load.mockImplementation(() => {
      if (rejectExecuted) {
        throw new Error("GeneralException");
      }
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(rejectExecuted).toBe(true);
    expect(result.status).toBe("error");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.error).toBe("GeneralException");
  });

  it("signals the disable CTA when rejecting the final pending artifact reaches zero", async () => {
    const suggestion = makeSuggestion({
      id: "s-final",
      anchor: "texto original",
      suggestedText: "",
      context: "Contexto con texto original.",
    });
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-final",
      ccTitle: "texto original",
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

  it("returns error when rejecting tracked changes causes comment.delete() sync to throw GeneralException", async () => {
    // Reproduces the real-host bug: tc.reject() invalidates the Word context, so
    // the context.sync() inside deleteLocatedStylisticComment throws GeneralException
    // AFTER the tracked changes were already successfully rejected in Word.
    // Atomic contract: if cleanup explodes after tracked-change mutation, the
    // workflow MUST return error instead of lying with terminal success.
    const suggestion = makeSuggestion({
      id: "chunk0-3",
      anchor: "sería a quien lo desconocía",
      suggestedText: "sería quien lo desconocía",
      context: "sería a quien lo desconocía. Su vida cambiaría.",
    });

    const rejectSpy1 = vi.fn();
    const rejectSpy2 = vi.fn();
    let cleanupStarted = false;

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
          delete: vi.fn(() => {
            cleanupStarted = true;
          }),
        },
      ],
    });

    // Simulate: tracked changes reject successfully, but the later comment cleanup
    // sync explodes with GeneralException after semantic resolution already happened.
    context.sync.mockImplementation(() => {
      if (cleanupStarted) {
        return Promise.reject(new Error("GeneralException"));
      }
      return Promise.resolve();
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("error");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.error).toBe("GeneralException");
    expect(rejectSpy1).toHaveBeenCalledOnce();
    expect(rejectSpy2).toHaveBeenCalledOnce();
  });

  it("returns error in one click when immediate re-observation finishes the remaining replace side but cleanup fails", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-4",
      anchor: "sándwich o sánduche",
      suggestedText: "sándwich o sánguche",
      context: "Quería debatir si debía decir sándwich o sánduche antes de hacer el pedido.",
    });

    let resolutionPhase = 0;
    let cleanupAttemptedAfterFullResolution = false;
    let shouldThrowOnAddedReject = true;

    const deletedRejectSpy = vi.fn(() => {
      resolutionPhase = Math.max(resolutionPhase, 1);
    });
    const addedRejectSpy = vi.fn(() => {
      if (shouldThrowOnAddedReject) {
        shouldThrowOnAddedReject = false;
        throw new Error("ItemNotFound");
      }

      resolutionPhase = 2;
    });
    const commentDeleteSpy = vi.fn(() => {
      cleanupAttemptedAfterFullResolution = true;
    });

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
      ccTag: "stylistic:track-change:chunk0-4",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-4",
        insertedTag: "stylistic:track-change:chunk0-4",
        deletedValue: "sándwich o sánduche",
        anchorValue:
          "Quería debatir si debía decir sándwich o sánduche antes de hacer el pedido.",
      }),
      comments: [comment],
    });

    const buildTrackedChanges = () => {
      if (resolutionPhase === 0) {
        return [
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
        ];
      }

      if (resolutionPhase === 1) {
        return [
          {
            id: "tc-added",
            type: "Added",
            accept: vi.fn(),
            reject: addedRejectSpy,
          },
        ];
      }

      return [];
    };

    context._cc.getTrackedChanges.mockImplementation(() => ({
      items: buildTrackedChanges(),
      load: vi.fn(),
    }));
    const getCcRange = context._cc.getRange as unknown as () => {
      getTrackedChanges: ReturnType<typeof vi.fn>;
    };
    const ccRange = getCcRange();
    ccRange.getTrackedChanges.mockImplementation(() => ({
      items: [],
      load: vi.fn(),
    }));
    context.document.body.getTrackedChanges.mockImplementation(() => ({
      items: [],
      load: vi.fn(),
    }));

    context.sync.mockImplementation(() => {
      if (cleanupAttemptedAfterFullResolution) {
        return Promise.reject(new Error("ItemNotFound"));
      }
      return Promise.resolve();
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("error");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.commentDeleted).toBe(false);
    expect(result.errorPhase).toBe("cleanup-comment");
    expect(result.executionReport).toEqual({
      attempted: 2,
      completed: 2,
      remaining: 0,
    });
    expect(result.error).toBe("ItemNotFound");
    expect(addedRejectSpy).toHaveBeenCalledTimes(2);
    expect(deletedRejectSpy).toHaveBeenCalledOnce();
    expect(commentDeleteSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).not.toHaveBeenCalled();
  });

  it("re-observes the remaining replace side with fresh proxies after the first reject sync", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-fresh-reobserve-reject",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let phase = 0;
    const callOrder: string[] = [];

    const deletedRejectInitial = vi.fn(() => {
      callOrder.push("reject-deleted-initial");
      phase = 1;
    });
    const addedRejectStale = vi.fn(() => {
      callOrder.push("reject-added-stale");
      throw new Error("ItemNotFound");
    });
    const addedRejectFresh = vi.fn(() => {
      callOrder.push("reject-added-fresh");
      phase = 2;
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-fresh-reobserve-reject",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-fresh-reobserve-reject",
        insertedTag: "stylistic:track-change:chunk0-fresh-reobserve-reject",
        deletedValue: "ni Shu",
        anchorValue: "No sabían si venía de ni Shu o de otro sitio.",
      }),
      comments: [],
    });

    context._cc.getTrackedChanges.mockImplementation(() => {
      if (phase === 0) {
        return {
          items: [
            {
              id: "tc-added-stale",
              type: "Added",
              accept: vi.fn(),
              reject: addedRejectStale,
            },
            {
              id: "tc-deleted-initial",
              type: "Deleted",
              accept: vi.fn(),
              reject: deletedRejectInitial,
            },
          ],
          load: vi.fn(),
        };
      }

      if (phase === 1) {
        return {
          items: [
            {
              id: "tc-added-fresh",
              type: "Added",
              accept: vi.fn(),
              reject: addedRejectFresh,
            },
          ],
          load: vi.fn(),
        };
      }

      return {
        items: [],
        load: vi.fn(),
      };
    });

    const getCcRange = context._cc.getRange as unknown as () => {
      getTrackedChanges: ReturnType<typeof vi.fn>;
    };
    const ccRange = getCcRange();
    ccRange.getTrackedChanges.mockImplementation(() => ({
      items: [],
      load: vi.fn(),
    }));
    context.document.body.getTrackedChanges.mockImplementation(() => ({
      items: [],
      load: vi.fn(),
    }));

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder).toEqual([
      "reject-deleted-initial",
      "reject-added-fresh",
    ]);
    expect(addedRejectStale).not.toHaveBeenCalled();
    expect(addedRejectFresh).toHaveBeenCalledOnce();
  });

  it("re-resolves the preferred candidate to a fresh proxy instead of reusing a stale content control during reject re-observation", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-preferred-fresh-proxy-reject",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let phase = 0;
    let getByTagCalls = 0;
    const callOrder: string[] = [];

    const deletedRejectInitial = vi.fn(() => {
      callOrder.push("reject-deleted-initial");
      phase = 1;
    });
    const addedRejectInitial = vi.fn(() => {
      callOrder.push("reject-added-initial-should-not-run");
      throw new Error("stale-added-should-not-run");
    });
    const addedRejectFresh = vi.fn(() => {
      callOrder.push("reject-added-fresh");
      phase = 2;
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-preferred-fresh-proxy-reject",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-preferred-fresh-proxy-reject",
        insertedTag:
          "stylistic:track-change:chunk0-preferred-fresh-proxy-reject",
        deletedValue: "ni Shu",
        anchorValue: "No sabían si venía de ni Shu o de otro sitio.",
      }),
      comments: [],
    });

    context._cc.getTrackedChanges.mockImplementation(() => ({
      items:
        phase === 0
          ? [
              {
                id: "tc-added-initial",
                type: "Added",
                accept: vi.fn(),
                reject: addedRejectInitial,
              },
              {
                id: "tc-deleted-initial",
                type: "Deleted",
                accept: vi.fn(),
                reject: deletedRejectInitial,
              },
            ]
          : [],
      load: vi.fn(),
    }));

    const initialCcRange = (context._cc.getRange as unknown as () => {
      getTrackedChanges: ReturnType<typeof vi.fn>;
    })();
    initialCcRange.getTrackedChanges.mockImplementation(() => ({
      items: [],
      load: vi.fn(),
    }));

    context._cc.load.mockImplementation(() => {
      if (phase >= 1) {
        throw new Error("stale-preferred-cc");
      }
    });

    const freshRangeTCCollection = { load: vi.fn() };
    const getFreshAddedRangeItems = () => {
      if (phase !== 1) {
        return [];
      }

      return [
        {
          id: "tc-added-fresh",
          type: "Added",
          accept: vi.fn(),
          reject: addedRejectFresh,
        },
      ];
    };
    const freshCc = {
      title: context._cc.title,
      tag: context._cc.tag,
      load: vi.fn(),
      getTrackedChanges: vi.fn(() => ({ items: [], load: vi.fn() })),
      getRange: vi.fn(() => ({
        compareLocationWith: vi.fn(() => ({ value: "Equal" })),
        getTrackedChanges: vi.fn(() => ({
          ...freshRangeTCCollection,
          items: getFreshAddedRangeItems(),
        })),
      })),
      delete: vi.fn(),
    };

    context.document.body.getTrackedChanges.mockImplementation(() => ({
      items: [],
      load: vi.fn(),
    }));

    context.document.contentControls.getByTag.mockImplementation(() => {
      getByTagCalls += 1;
      return {
        items: getByTagCalls === 1 ? [context._cc] : [freshCc],
        load: vi.fn(),
      };
    });

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.error).toBeUndefined();
    expect(callOrder).toEqual([
      "reject-deleted-initial",
      "reject-added-fresh",
    ]);
    expect(addedRejectInitial).not.toHaveBeenCalled();
    expect(addedRejectFresh).toHaveBeenCalledOnce();
  });

  it("re-observes the remaining Added side without requiring a fresh Deleted pair", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-side-specific-deleted",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let phase = 0;
    const callOrder: string[] = [];

    const deletedRejectInitial = vi.fn(() => {
      callOrder.push("reject-deleted-initial");
      phase = 1;
    });
    const addedRejectFresh = vi.fn(() => {
      callOrder.push("reject-added-fresh");
      phase = 2;
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-side-specific-deleted",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-side-specific-deleted",
        insertedTag: "stylistic:track-change:chunk0-side-specific-deleted",
        deletedValue: "ni Shu",
        anchorValue: "No sabían si venía de ni Shu o de otro sitio.",
      }),
      comments: [],
    });

    context._cc.getTrackedChanges.mockImplementation(() => ({
      items:
        phase === 0
          ? [
              {
                id: "tc-added-initial",
                type: "Added",
                accept: vi.fn(),
                reject: vi.fn(() => {
                  throw new Error("stale-added-should-not-run");
                }),
              },
              {
                id: "tc-deleted-initial",
                type: "Deleted",
                accept: vi.fn(),
                reject: deletedRejectInitial,
              },
            ]
          : [],
      load: vi.fn(),
    }));

    const getCcRange = context._cc.getRange as unknown as () => {
      getTrackedChanges: ReturnType<typeof vi.fn>;
    };
    const ccRange = getCcRange();
    ccRange.getTrackedChanges.mockImplementation(() => {
      if (phase === 1) {
        return {
          items: [
            {
              id: "tc-added-fresh",
              type: "Added",
              accept: vi.fn(),
              reject: addedRejectFresh,
            },
          ],
          load: vi.fn(),
        };
      }

      return {
        items: [],
        load: vi.fn(),
      };
    });
    context.document.body.getTrackedChanges.mockImplementation(() => ({
      items: [],
      load: vi.fn(),
    }));

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("rejected");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder).toEqual([
      "reject-deleted-initial",
      "reject-added-fresh",
    ]);
  });

  it("returns error and skips cleanup when reject leaves one semantic side pending after execution", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-half-after-reject",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let phase = 0;
    const addedRejectSpy = vi.fn(() => {
      phase = 1;
    });
    const deletedRejectSpy = vi.fn(() => {
      phase = 2;
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-half-after-reject",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-half-after-reject",
        insertedTag: "stylistic:track-change:chunk0-half-after-reject",
        deletedValue: "ni Shu",
        anchorValue: "No sabían si venía de ni Shu o de otro sitio.",
      }),
      comments: [],
    });

    context._cc.getTrackedChanges.mockImplementation(() => ({
      items:
        phase === 0
          ? [
              {
                id: "tc-added-initial",
                type: "Added",
                accept: vi.fn(),
                reject: addedRejectSpy,
              },
              {
                id: "tc-deleted-initial",
                type: "Deleted",
                accept: vi.fn(),
                reject: deletedRejectSpy,
              },
            ]
          : [],
      load: vi.fn(),
    }));

    const getCcRange = context._cc.getRange as unknown as () => {
      getTrackedChanges: ReturnType<typeof vi.fn>;
    };
    const ccRange = getCcRange();
    ccRange.getTrackedChanges.mockImplementation(() => {
      if (phase === 1) {
        return {
          items: [
            {
              id: "tc-deleted-fresh",
              type: "Deleted",
              accept: vi.fn(),
              reject: deletedRejectSpy,
            },
          ],
          load: vi.fn(),
        };
      }

      if (phase >= 2) {
        return {
          items: [
            {
              id: "tc-deleted-still-pending",
              type: "Deleted",
              accept: vi.fn(),
              reject: vi.fn(),
            },
          ],
          load: vi.fn(),
        };
      }

      return {
        items: [],
        load: vi.fn(),
      };
    });
    context.document.body.getTrackedChanges.mockImplementation(() => ({
      items: [],
      load: vi.fn(),
    }));

    installWordWithContext(context);

    const result = await adapter.rejectSuggestion(suggestion);

    expect(result.status).toBe("error");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.error).toContain("falso success");
    expect(context._cc.delete).not.toHaveBeenCalled();
  });
});
