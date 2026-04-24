import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "./WordAdapter";
import type { ITelemetryPort } from "../../domain/ports";
import {
  installRejectingWord,
  makeCompoundV2Title,
  installWordWithContext,
  makeResolveSuggestionContext,
  makeSuggestion,
} from "./WordAdapterActionTestHelper";

type TestTrackedChange = {
  accept?: () => unknown;
};

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

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.commentDeleted).toBe(true);
    expect(tcAccept1).toHaveBeenCalledOnce();
    expect(tcAccept2).toHaveBeenCalledOnce();
    expect(commentDeleteSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("accepts replace tracked changes by confirming insertion before deleting the original", async () => {
    const suggestion = makeSuggestion({
      id: "s-ordered-accept",
      anchor: "desde allí",
      suggestedText: "desde entonces",
      context:
        "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
    });
    const callOrder: string[] = [];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-ordered-accept",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-ordered-accept",
        insertedTag: "stylistic:track-change:s-ordered-accept",
        deletedValue: "desde allí",
        anchorValue:
          "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
      }),
      spanTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(() => {
            callOrder.push("accept-added");
          }),
          reject: vi.fn(),
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(() => {
            callOrder.push("accept-deleted");
          }),
          reject: vi.fn(),
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result).toEqual(
      expect.objectContaining({
        status: "accepted",
      }),
    );

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder).toEqual(["accept-added", "accept-deleted"]);
  });

  it("syncs after each replace accept step before continuing with the next semantic side", async () => {
    const suggestion = makeSuggestion({
      id: "s-sync-accept",
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
      ccTag: "stylistic:track-change:s-sync-accept",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-sync-accept",
        insertedTag: "stylistic:track-change:s-sync-accept",
        deletedValue: "desde allí",
        anchorValue:
          "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
      }),
      spanTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(() => {
            queuedStep = 1;
            callOrder.push("accept-added");
          }),
          reject: vi.fn(),
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(() => {
            queuedStep = 2;
            callOrder.push("accept-deleted");
          }),
          reject: vi.fn(),
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

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder.slice(0, 4)).toEqual([
      "accept-added",
      "sync",
      "accept-deleted",
      "sync",
    ]);
  });

  it("returns cc-not-found when the Content Control anchor is missing", async () => {
    const suggestion = makeSuggestion({
      anchor: "texto original",
      suggestedText: "",
      context: "Contexto con texto original.",
    });

    const context = makeResolveSuggestionContext({ ccFound: false });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("cc-not-found");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.commentDeleted).toBe(false);
  });

  it("returns unobservable when the CC remains but no tracked changes can be observed", async () => {
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

    expect(result.status).toBe("unobservable");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.error).toContain("Word no expuso suficientes tracked changes");
    expect(context._cc.delete).not.toHaveBeenCalled();
  });

  it("accepts compound-v2 replace suggestions when rich metadata and tracked changes are present", async () => {
    const suggestion = makeSuggestion({
      id: "s-1",
      anchor: "texto original",
      suggestedText: "texto sugerido",
      context: "Contexto con texto original.",
    });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTitle: makeCompoundV2Title(),
      spanTCItems: [
        { id: "tc-added", type: "Added", accept: addedAcceptSpy, reject: vi.fn() },
        { id: "tc-deleted", type: "Deleted", accept: deletedAcceptSpy, reject: vi.fn() },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("returns identity-lost when compound-v2 metadata is structurally corrupt or incomplete", async () => {
    const suggestion = makeSuggestion({ id: "s-1" });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTitle: makeCompoundV2Title({ overrides: { deletedSideRef: undefined } }),
      spanTCItems: [],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("identity-lost");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.error).toContain("compound-v2");
    expect(context._cc.delete).not.toHaveBeenCalled();
  });

  it("treats compound-v2 replace suggestions without visible tracked changes as unobservable", async () => {
    const suggestion = makeSuggestion({ id: "s-1" });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTitle: makeCompoundV2Title(),
      spanTCItems: [],
      bodyTCItems: [],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("unobservable");
  });

  it("does not resolve bare-id artifacts once compound-v2 becomes mandatory", async () => {
    const suggestion = makeSuggestion({ id: "legacy-1" });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "legacy-1",
      ccTitle: "texto original",
      spanTCItems: [{ type: "Deleted", accept: vi.fn(), reject: vi.fn() }],
      comments: [],
    });
    context.document.contentControls.getByTag = vi.fn(() => ({
      items: [],
      load: vi.fn(),
    }));
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("cc-not-found");
  });

  it("accepts tracked changes even when the associated comment is already gone", async () => {
    const suggestion = makeSuggestion({
      anchor: "texto original",
      suggestedText: "",
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
      ccTitle: "texto original",
      spanTCItems,
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(1);
    expect(result.commentDeleted).toBe(false);
    expect(result.pendingAfter.hasPendingStylisticArtifacts).toBe(true);
    expect(result.error).toBeUndefined();
    expect(tcAcceptSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("returns unobservable when a replace suggestion exposes only the inserted side", async () => {
    const suggestion = makeSuggestion({
      id: "s-half-visible",
      anchor: "desde allí",
      suggestedText: "desde entonces",
      context:
        "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
    });
    const addedAcceptSpy = vi.fn();

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
          accept: addedAcceptSpy,
          reject: vi.fn(),
        },
      ],
      rangeTCItems: [],
      bodyTCItems: [],
      comments: [],
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("unobservable");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.commentDeleted).toBe(false);
    expect(result.error).toContain("Word no expuso suficientes tracked changes");
    expect(addedAcceptSpy).not.toHaveBeenCalled();
    expect(context._cc.delete).not.toHaveBeenCalled();
  });

  it("accepts and deletes the matching comment by content when authorName is not Stylistic", async () => {
    const suggestion = makeSuggestion({
      anchor: "texto original",
      suggestedText: "",
      context: "Contexto con texto original.",
    });
    const tcAcceptSpy = vi.fn();
    const targetCommentDeleteSpy = vi.fn();
    const secondCommentDeleteSpy = vi.fn();
    const commentRange = {
      compareLocationWith: vi.fn(() => ({ value: "Equal" })),
    };

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-1",
      spanTCItems: [
        {
          type: "Deleted",
          accept: tcAcceptSpy,
          reject: vi.fn(),
        },
      ],
      ccTitle: "texto original",
      comments: [
        {
          authorName: "Usuario de prueba",
          content: "[Claridad]\nMas claro",
          getRange: vi.fn(() => commentRange),
          delete: targetCommentDeleteSpy,
        },
        {
          authorName: "Usuario de prueba",
          content: "[Registro]\nOtra cosa",
          getRange: vi.fn(() => commentRange),
          delete: secondCommentDeleteSpy,
        },
      ],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.commentDeleted).toBe(true);
    expect(tcAcceptSpy).toHaveBeenCalledOnce();
    expect(targetCommentDeleteSpy).toHaveBeenCalledOnce();
    expect(secondCommentDeleteSpy).not.toHaveBeenCalled();
  });

  it("accepts and deletes the first colocated Stylistic comment even when content differs", async () => {
    const suggestion = makeSuggestion({
      anchor: "texto original",
      suggestedText: "",
      context: "Contexto con texto original.",
      category: "Gramática",
      justification: "Ajuste actualizado",
    });
    const tcAcceptSpy = vi.fn();
    const colocatedDeleteSpy = vi.fn();
    const distantDeleteSpy = vi.fn();
    const colocatedRange = {
      compareLocationWith: vi.fn(() => ({ value: "Equal" })),
    };
    const distantRange = {
      compareLocationWith: vi.fn(() => ({ value: "Before" })),
    };

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-1",
      spanTCItems: [
        {
          type: "Deleted",
          accept: tcAcceptSpy,
          reject: vi.fn(),
        },
      ],
      ccTitle: "texto original",
      comments: [
        {
          authorName: "Usuario de prueba",
          content: "[Gramática]\nVersión previa del comentario",
          getRange: vi.fn(() => colocatedRange),
          delete: colocatedDeleteSpy,
        },
        {
          authorName: "Usuario de prueba",
          content: "[Claridad]\nComentario de otra sugerencia",
          getRange: vi.fn(() => distantRange),
          delete: distantDeleteSpy,
        },
      ],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.commentDeleted).toBe(true);
    expect(tcAcceptSpy).toHaveBeenCalledOnce();
    expect(colocatedDeleteSpy).toHaveBeenCalledOnce();
    expect(distantDeleteSpy).not.toHaveBeenCalled();
  });

  it("accepts and deletes a colocated Stylistic comment when Word returns CRLF content", async () => {
    const suggestion = makeSuggestion({
      anchor: "texto original",
      suggestedText: "",
      context: "Contexto con texto original.",
      category: "Gramática",
      justification: "Ajuste actualizado",
    });
    const tcAcceptSpy = vi.fn();
    const commentDeleteSpy = vi.fn();
    const commentRange = {
      compareLocationWith: vi.fn(() => ({ value: "Equal" })),
    };

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-1",
      spanTCItems: [
        {
          type: "Deleted",
          accept: tcAcceptSpy,
          reject: vi.fn(),
        },
      ],
      ccTitle: "texto original",
      comments: [
        {
          authorName: "Usuario de prueba",
          content: "[Gramática]\r\nComentario desde Word",
          getRange: vi.fn(() => commentRange),
          delete: commentDeleteSpy,
        },
      ],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.commentDeleted).toBe(true);
    expect(tcAcceptSpy).toHaveBeenCalledOnce();
    expect(commentDeleteSpy).toHaveBeenCalledOnce();
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

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("does not over-collect replace tracked changes when multiple evidence sources expose stale extras", async () => {
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
          accept: vi.fn(() => callOrder.push("accept-added-main")),
          reject: vi.fn(),
        },
      ],
      rangeTCItems: [],
      bodyTCItems: [
        {
          id: "tc-added-main",
          type: "Added",
          accept: vi.fn(() => callOrder.push("accept-added-duplicate")),
          reject: vi.fn(),
        },
        {
          id: "tc-deleted-stale-body",
          type: "Deleted",
          accept: vi.fn(() => callOrder.push("accept-deleted-stale-body")),
          reject: vi.fn(),
        },
      ],
      bodyTCRelations: ["Equal", "AdjacentBefore"],
      deletedSideText: "ni Shu",
      deletedSideRangeTCItems: [
        {
          id: "tc-deleted-main",
          type: "Deleted",
          accept: vi.fn(() => callOrder.push("accept-deleted-main")),
          reject: vi.fn(),
        },
      ],
      operationalAnchorText: "No sabían si venía de ni Shu o de otro sitio.",
      operationalAnchorRangeTCItems: [
        {
          id: "tc-added-stale-anchor",
          type: "Added",
          accept: vi.fn(() => callOrder.push("accept-added-stale-anchor")),
          reject: vi.fn(),
        },
        {
          id: "tc-deleted-stale-anchor",
          type: "Deleted",
          accept: vi.fn(() => callOrder.push("accept-deleted-stale-anchor")),
          reject: vi.fn(),
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder).toEqual(["accept-added-main", "accept-deleted-stale-body"]);
  });

  it("normalizes duplicate Deleted evidence into one semantic replace pair after the Added side resolves first", async () => {
    const suggestion = makeSuggestion({
      id: "s-duplicate-deleted-evidence",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context:
        "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
    });
    const callOrder: string[] = [];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-duplicate-deleted-evidence",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-duplicate-deleted-evidence",
        insertedTag: "stylistic:track-change:s-duplicate-deleted-evidence",
        deletedValue: "ni Shu",
        anchorValue:
          "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
      }),
      spanTCItems: [
        {
          id: "tc-deleted-cc",
          type: "Deleted",
          accept: vi.fn(() => callOrder.push("accept-deleted-cc")),
          reject: vi.fn(),
        },
      ],
      rangeTCItems: [
        {
          id: "tc-deleted-range",
          type: "Deleted",
          accept: vi.fn(() => callOrder.push("accept-deleted-range")),
          reject: vi.fn(),
        },
      ],
      bodyTCItems: [
        {
          id: "tc-added-body",
          type: "Added",
          accept: vi.fn(() => callOrder.push("accept-added-body")),
          reject: vi.fn(),
        },
      ],
      bodyTCRelations: ["Equal"],
      comments: [],
    });
    let queuedTrackedActions = 0;
    let syncedTrackedActions = 0;
    let trackedChangeSyncs = 0;

    const getCcTrackedChanges = context._cc
      .getTrackedChanges as unknown as () => { items: TestTrackedChange[] };
    const getCcRange = context._cc.getRange as unknown as () => {
      getTrackedChanges: () => { items: TestTrackedChange[] };
    };
    const getBodyTrackedChanges = context.document.body
      .getTrackedChanges as unknown as () => { items: TestTrackedChange[] };

    for (const trackedChange of [
      ...getCcTrackedChanges().items,
      ...getCcRange().getTrackedChanges().items,
      ...getBodyTrackedChanges().items,
    ]) {
      const originalAccept = trackedChange.accept;
      trackedChange.accept = vi.fn(() => {
        queuedTrackedActions += 1;
        return originalAccept?.();
      });
    }

    context.sync.mockImplementation(async () => {
      if (queuedTrackedActions > syncedTrackedActions) {
        syncedTrackedActions = queuedTrackedActions;
        trackedChangeSyncs += 1;
      }
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(trackedChangeSyncs).toBe(2);
    // ccRange+bodyRelated is now the primary evidence surface: it targets
    // the real document-level tracked-change proxies instead of the
    // cc-internal ones that Word can silently no-op on.
    expect(callOrder).toEqual(["accept-added-body", "accept-deleted-range"]);
  });

  it("prefers CC and body replace evidence before deleted-side fallback when duplicate deleted proxies exist", async () => {
    const suggestion = makeSuggestion({
      id: "s-prefer-baseline-replace-evidence",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context:
        "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
    });
    let usedBodyAdded = false;
    let usedDeletedSideAdded = false;

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-prefer-baseline-replace-evidence",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-prefer-baseline-replace-evidence",
        insertedTag: "stylistic:track-change:s-prefer-baseline-replace-evidence",
        deletedValue: "ni Shu",
        anchorValue:
          "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
      }),
      spanTCItems: [
        {
          id: "tc-deleted-cc",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(),
        },
      ],
      rangeTCItems: [
        {
          id: "tc-deleted-range",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(),
        },
      ],
      bodyTCItems: [
        {
          id: "tc-added-body",
          type: "Added",
          accept: vi.fn(() => {
            usedBodyAdded = true;
          }),
          reject: vi.fn(),
        },
      ],
      bodyTCRelations: ["Equal"],
      deletedSideText: "ni Shu",
      deletedSideRangeTCItems: [
        {
          id: "tc-added-deleted-side",
          type: "Added",
          accept: vi.fn(() => {
            usedDeletedSideAdded = true;
          }),
          reject: vi.fn(),
        },
      ],
      comments: [],
    });

    context.sync.mockImplementation(async () => {
      if (usedDeletedSideAdded) {
        throw new Error("ItemNotFound");
      }
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(usedBodyAdded).toBe(true);
    expect(usedDeletedSideAdded).toBe(false);
  });

  it("prefers the deletedSide Deleted proxy over the CC Deleted proxy when the duplicated-side path fires and deletedSide exposes a real Deleted TC", async () => {
    // Regression: when cc.getTrackedChanges() returns an internal CC change (type Deleted)
    // and ccRange also returns a Deleted, the early-path previously returned cc+ccRange+bodyRelated
    // and normalization picked the cc Deleted (a spurious intra-CC change), leaving the actual
    // replace-pair deletion at the deletedSide location untouched.
    const suggestion = makeSuggestion({
      id: "s-prefer-deleted-side-deleted",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context:
        "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
    });
    const usedCcDeleted: string[] = [];
    const usedDeletedSideDeleted: string[] = [];
    const usedBodyAdded: string[] = [];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-prefer-deleted-side-deleted",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-prefer-deleted-side-deleted",
        insertedTag: "stylistic:track-change:s-prefer-deleted-side-deleted",
        deletedValue: "ni Shu",
        anchorValue:
          "Xia no tenía idea de lo que estaba pasando por la mente de Mei ni Shu.",
      }),
      spanTCItems: [
        {
          id: "tc-deleted-cc-internal",
          type: "Deleted",
          accept: vi.fn(() => {
            usedCcDeleted.push("cc-internal");
          }),
          reject: vi.fn(),
        },
      ],
      rangeTCItems: [
        {
          id: "tc-deleted-cc-range",
          type: "Deleted",
          accept: vi.fn(() => {
            usedCcDeleted.push("cc-range");
          }),
          reject: vi.fn(),
        },
      ],
      bodyTCItems: [
        {
          id: "tc-added-body",
          type: "Added",
          accept: vi.fn(() => {
            usedBodyAdded.push("body");
          }),
          reject: vi.fn(),
        },
      ],
      bodyTCRelations: ["Equal"],
      deletedSideText: "ni Shu",
      deletedSideRangeTCItems: [
        {
          id: "tc-deleted-actual",
          type: "Deleted",
          accept: vi.fn(() => {
            usedDeletedSideDeleted.push("deleted-side");
          }),
          reject: vi.fn(),
        },
      ],
      comments: [],
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    // ccRange+bodyRelated is the primary source, so the cc-range Deleted and
    // the body Added are chosen. The spurious cc-internal Deleted is never
    // touched (it is cc.getTrackedChanges() only, which is deprioritized),
    // and the deletedSide fallback is never needed because the primary pair
    // already completes at the document-level scope.
    expect(usedCcDeleted).toEqual(["cc-range"]);
    expect(usedBodyAdded).toEqual(["body"]);
    expect(usedDeletedSideDeleted).toEqual([]);
  });

  it("normalizes duplicate deleted-side observations into a single semantic replace pair", async () => {
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
          accept: vi.fn(() => callOrder.push("accept-deleted-cc")),
          reject: vi.fn(),
        },
      ],
      rangeTCItems: [
        {
          id: "tc-added-range",
          type: "Added",
          accept: vi.fn(() => callOrder.push("accept-added-range")),
          reject: vi.fn(),
        },
      ],
      bodyTCItems: [
        {
          id: "tc-deleted-body",
          type: "Deleted",
          accept: vi.fn(() => callOrder.push("accept-deleted-body")),
          reject: vi.fn(),
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
            accept: vi.fn(() => callOrder.push("accept-deleted-comment")),
            reject: vi.fn(),
          },
        ],
      ],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    // ccRange+bodyRelated is primary: range exposes the Added, bodyRelated
    // (AdjacentBefore) exposes the Deleted. The cc-internal Deleted is skipped.
    expect(callOrder).toEqual(["accept-added-range", "accept-deleted-body"]);
  });

  it("accepts replace suggestions when only the CC range exposes tracked changes in real-host style semantics", async () => {
    const suggestion = makeSuggestion({
      anchor: "quién",
      suggestedText: "quien",
      context: "Contexto con quién.",
    });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

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
          accept: addedAcceptSpy,
          reject: vi.fn(),
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: deletedAcceptSpy,
          reject: vi.fn(),
        },
      ],
      bodyTCItems: [],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("accepts replace suggestions when only the operational anchor range exposes tracked changes", async () => {
    const suggestion = makeSuggestion({
      id: "s-anchor-only",
      anchor: "parecía ser",
      suggestedText: "parece que es",
      context: "Hasta donde sabía, parecía ser la única al tanto.",
    });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

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
          accept: addedAcceptSpy,
          reject: vi.fn(),
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: deletedAcceptSpy,
          reject: vi.fn(),
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("accepts replace suggestions when the deleted-side locator exposes the missing deleted tracked change", async () => {
    const suggestion = makeSuggestion({
      id: "s-pre-replace-context",
      anchor: "parecía ser",
      suggestedText: "parece que es",
      context: "Hasta donde sabía, parecía ser la única al tanto.",
    });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

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
          accept: addedAcceptSpy,
          reject: vi.fn(),
        },
      ],
      rangeTCItems: [],
      bodyTCItems: [],
      deletedSideText: "parecía ser",
      deletedSideRangeTCItems: [
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: deletedAcceptSpy,
          reject: vi.fn(),
        },
      ],
      operationalAnchorText: undefined,
      operationalAnchorRangeTCItems: [],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("accepts replace suggestions when only the colocated comment range exposes tracked changes", async () => {
    const suggestion = makeSuggestion({
      id: "s-comment-range",
      anchor: "parecía ser",
      suggestedText: "parece que es",
      context: "Hasta donde sabía, parecía ser la única al tanto.",
      category: "Gramática",
      justification: "Ajuste verbal",
    });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

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
            accept: addedAcceptSpy,
            reject: vi.fn(),
          },
          {
            id: "tc-deleted",
            type: "Deleted",
            accept: deletedAcceptSpy,
            reject: vi.fn(),
          },
        ],
      ],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("prefers the compound-v2 CC when multiple CCs share the same tag", async () => {
    const suggestion = makeSuggestion({ id: "chunk0-0" });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

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
              accept: addedAcceptSpy,
              reject: vi.fn(),
            },
            {
              id: "tc-deleted",
              type: "Deleted",
              accept: deletedAcceptSpy,
              reject: vi.fn(),
            },
          ],
          rangeTCItems: [],
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("ignores stale compound-v2 CCs whose anchor metadata belongs to an earlier run", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-0",
      anchor: "fragmento actual",
      context: "Contexto con fragmento actual.",
    });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-0",
      ccItems: [
        {
          tag: "stylistic:track-change:chunk0-0",
          title: makeCompoundV2Title({
            suggestionId: "chunk0-0",
            insertedTag: "stylistic:track-change:chunk0-0",
            deletedValue: "anchor viejo",
            anchorValue: "Contexto viejo.",
          }),
          spanTCItems: [],
          rangeTCItems: [],
        },
        {
          tag: "stylistic:track-change:chunk0-0",
          title: makeCompoundV2Title({
            suggestionId: "chunk0-0",
            insertedTag: "stylistic:track-change:chunk0-0",
            deletedValue: "fragmento actual",
            anchorValue: "Contexto con fragmento actual.",
          }),
          spanTCItems: [
            {
              id: "tc-added",
              type: "Added",
              accept: addedAcceptSpy,
              reject: vi.fn(),
            },
            {
              id: "tc-deleted",
              type: "Deleted",
              accept: deletedAcceptSpy,
              reject: vi.fn(),
            },
          ],
          rangeTCItems: [],
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("still resolves when the only compound-v2 artifact has host-drifted deleted and anchor text", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-0",
      anchor: "fragmento actual",
      context: "Contexto con fragmento actual.",
    });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-0",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-0",
        insertedTag: "stylistic:track-change:chunk0-0",
        deletedValue: "anchor viejo que Word preservó",
        anchorValue: "Contexto viejo rehidratado por Word.",
      }),
      spanTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: addedAcceptSpy,
          reject: vi.fn(),
        },
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: deletedAcceptSpy,
          reject: vi.fn(),
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("tries later CC candidates when the first v2 candidate remains unobservable", async () => {
    const suggestion = makeSuggestion({ id: "chunk0-0" });
    const addedAcceptSpy = vi.fn();
    const deletedAcceptSpy = vi.fn();

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
              accept: addedAcceptSpy,
              reject: vi.fn(),
            },
            {
              id: "tc-deleted",
              type: "Deleted",
              accept: deletedAcceptSpy,
              reject: vi.fn(),
            },
          ],
          rangeTCItems: [],
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("accepts cross-sentence replace suggestions when the operational anchor exposes a missing deleted fragment", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-1",
      anchor: "multitud. Así que",
      suggestedText: "multitud, así que",
      context:
        "pero el flujo era muy bajo para considerarse una multitud. Así que Xia y Shu no tenían que dar explicaciones",
    });
    const addedAcceptSpy = vi.fn();
    const deletedTailAcceptSpy = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-1",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-1",
        insertedTag: "stylistic:track-change:chunk0-1",
        deletedValue: "multitud. Así que",
        anchorValue:
          "pero el flujo era muy bajo para considerarse una multitud. Así que Xia y Shu no tenían que dar explicaciones",
      }),
      spanTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: addedAcceptSpy,
          reject: vi.fn(),
        },
      ],
      rangeTCItems: [],
      bodyTCItems: [],
      operationalAnchorText:
        "pero el flujo era muy bajo para considerarse una multitud. Así que Xia y Shu no tenían que dar explicaciones",
      operationalAnchorRangeTCItems: [
        {
          id: "tc-deleted-tail",
          type: "Deleted",
          accept: deletedTailAcceptSpy,
          reject: vi.fn(),
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedTailAcceptSpy).toHaveBeenCalledOnce();
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

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(0);
    expect(result.commentDeleted).toBe(true);
    expect(commentDeleteSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("signals the disable CTA when accepting the final pending artifact reaches zero", async () => {
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
          accept: vi.fn(() => {
            context.document.contentControls.items = [];
          }),
          reject: vi.fn(),
        },
      ],
      comments: [],
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.pendingAfter.pendingStylisticArtifacts).toBe(0);
    expect(result.documentState).toBe("ready-to-disable-track-changes");
  });

  it("returns error in one click when immediate re-observation finishes the remaining replace side but comment cleanup fails", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-4",
      anchor: "sándwich o sánduche",
      suggestedText: "sándwich o sánguche",
      context: "Quería debatir si debía decir sándwich o sánduche antes de hacer el pedido.",
    });

    let resolutionPhase = 0;
    let cleanupAttemptedAfterFullResolution = false;
    let shouldThrowOnAddedAccept = true;

    const addedAcceptSpy = vi.fn(() => {
      if (shouldThrowOnAddedAccept) {
        shouldThrowOnAddedAccept = false;
        throw new Error("ItemNotFound");
      }

      resolutionPhase = 2;
    });
    const deletedAcceptSpy = vi.fn(() => {
      resolutionPhase = Math.max(resolutionPhase, 1);
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
            id: "tc-deleted",
            type: "Deleted",
            accept: deletedAcceptSpy,
            reject: vi.fn(),
          },
          {
            id: "tc-added",
            type: "Added",
            accept: addedAcceptSpy,
            reject: vi.fn(),
          },
        ];
      }

      if (resolutionPhase === 1) {
        return [
          {
            id: "tc-added",
            type: "Added",
            accept: addedAcceptSpy,
            reject: vi.fn(),
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

    const result = await adapter.acceptSuggestion(suggestion);

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
    expect(addedAcceptSpy).toHaveBeenCalledTimes(2);
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
    expect(commentDeleteSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).not.toHaveBeenCalled();
  });

  it("re-observes the remaining Deleted side with fresh proxies after the first sync", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-fresh-reobserve-accept",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let phase = 0;
    const callOrder: string[] = [];

    const addedAcceptInitial = vi.fn(() => {
      callOrder.push("accept-added-initial");
      phase = 1;
    });
    const deletedAcceptStale = vi.fn(() => {
      callOrder.push("accept-deleted-stale");
      throw new Error("ItemNotFound");
    });
    const deletedAcceptFresh = vi.fn(() => {
      callOrder.push("accept-deleted-fresh");
      phase = 2;
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-fresh-reobserve-accept",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-fresh-reobserve-accept",
        insertedTag: "stylistic:track-change:chunk0-fresh-reobserve-accept",
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
                id: "tc-added-initial",
                type: "Added",
                accept: addedAcceptInitial,
                reject: vi.fn(),
              },
              {
                id: "tc-deleted-stale",
                type: "Deleted",
                accept: deletedAcceptStale,
                reject: vi.fn(),
              },
            ],
          load: vi.fn(),
        };
      }

      if (phase === 1) {
        return {
          items: [
            {
              id: "tc-deleted-fresh",
              type: "Deleted",
              accept: deletedAcceptFresh,
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

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder).toEqual([
      "accept-added-initial",
      "accept-deleted-fresh",
    ]);
    expect(deletedAcceptStale).not.toHaveBeenCalled();
    expect(deletedAcceptFresh).toHaveBeenCalledOnce();
  });

  it("falls back to one atomic accept batch when Word rejects the Added side alone but accepts the full replace together", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-atomic-accept-fallback",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    const queuedTypes: string[] = [];
    const callOrder: string[] = [];
    let resolutionPhase = 0;

    const deletedAcceptSpy = vi.fn(() => {
      queuedTypes.push("Deleted");
      callOrder.push(`accept-deleted-${queuedTypes.length}`);
    });
    const addedAcceptSpy = vi.fn(() => {
      queuedTypes.push("Added");
      callOrder.push(`accept-added-${queuedTypes.length}`);
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-atomic-accept-fallback",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-atomic-accept-fallback",
        insertedTag: "stylistic:track-change:chunk0-atomic-accept-fallback",
        deletedValue: "ni Shu",
        anchorValue: "No sabían si venía de ni Shu o de otro sitio.",
      }),
      comments: [],
    });

    context._cc.getTrackedChanges.mockImplementation(() => ({
      items:
        resolutionPhase === 0
          ? [
              {
                id: "tc-added",
                type: "Added",
                accept: addedAcceptSpy,
                reject: vi.fn(),
              },
              {
                id: "tc-deleted",
                type: "Deleted",
                accept: deletedAcceptSpy,
                reject: vi.fn(),
              },
            ]
          : [],
      load: vi.fn(),
    }));

    const getCcRange = context._cc.getRange as unknown as () => {
      getTrackedChanges: ReturnType<typeof vi.fn>;
    };
    const ccRange = getCcRange();
    ccRange.getTrackedChanges.mockImplementation(() => ({
      items:
        resolutionPhase === 0
          ? [
              {
                id: "tc-added",
                type: "Added",
                accept: addedAcceptSpy,
                reject: vi.fn(),
              },
              {
                id: "tc-deleted",
                type: "Deleted",
                accept: deletedAcceptSpy,
                reject: vi.fn(),
              },
            ]
          : [],
      load: vi.fn(),
    }));

    context.document.body.getTrackedChanges.mockImplementation(() => ({
      items: [],
      load: vi.fn(),
    }));

    context.sync.mockImplementation(async () => {
      const batchKey = queuedTypes.join(",");

      if (batchKey === "Added") {
        queuedTypes.length = 0;
        throw new Error("InvalidRibbonDefinition");
      }

      if (batchKey === "Added,Deleted") {
        queuedTypes.length = 0;
        resolutionPhase = 1;
        return;
      }

      queuedTypes.length = 0;
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.error).toBeUndefined();
    expect(callOrder).toEqual([
      "accept-added-1",
      "accept-added-1",
      "accept-deleted-2",
    ]);
    expect(addedAcceptSpy).toHaveBeenCalledTimes(2);
    expect(deletedAcceptSpy).toHaveBeenCalledOnce();
  });

  it("re-observes the remaining Deleted side without requiring a fresh Added pair", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-side-specific-added",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let phase = 0;
    const callOrder: string[] = [];

    const addedAcceptInitial = vi.fn(() => {
      callOrder.push("accept-added-initial");
      phase = 1;
    });
    const deletedAcceptFresh = vi.fn(() => {
      callOrder.push("accept-deleted-fresh");
      phase = 2;
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-side-specific-added",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-side-specific-added",
        insertedTag: "stylistic:track-change:chunk0-side-specific-added",
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
                accept: addedAcceptInitial,
                reject: vi.fn(),
              },
              {
                id: "tc-deleted-initial",
                type: "Deleted",
                accept: vi.fn(() => {
                  throw new Error("stale-deleted-should-not-run");
                }),
                reject: vi.fn(),
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
              accept: deletedAcceptFresh,
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

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder).toEqual([
      "accept-added-initial",
      "accept-deleted-fresh",
    ]);
  });

  it("ignores a stale Added reappearing from full observation after step 1 and keeps only the remaining Deleted side", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-ignore-stale-deleted-after-step1",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let phase = 0;
    const callOrder: string[] = [];

    const addedAcceptInitial = vi.fn(() => {
      callOrder.push("accept-added-initial");
      phase = 1;
    });
    const deletedAcceptFresh = vi.fn(() => {
      callOrder.push("accept-deleted-fresh");
      phase = 2;
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-ignore-stale-deleted-after-step1",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-ignore-stale-deleted-after-step1",
        insertedTag:
          "stylistic:track-change:chunk0-ignore-stale-deleted-after-step1",
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
                id: "tc-added-initial",
                type: "Added",
                accept: addedAcceptInitial,
                reject: vi.fn(),
              },
              {
                id: "tc-deleted-initial",
                type: "Deleted",
                accept: vi.fn(() => {
                  throw new Error("stale-deleted-should-not-run");
                }),
                reject: vi.fn(),
              },
            ],
          load: vi.fn(),
        };
      }

      if (phase === 1) {
        return {
          items: [
            {
              id: "tc-added-stale-after-step1",
              type: "Added",
              accept: vi.fn(() => {
                throw new Error("stale-added-should-not-run");
              }),
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

    const getCcRange = context._cc.getRange as unknown as () => {
      getTrackedChanges: ReturnType<typeof vi.fn>;
    };
    const ccRange = getCcRange();
    ccRange.getTrackedChanges.mockImplementation(() => ({
      items:
        phase === 1
          ? [
              {
                id: "tc-deleted-fresh",
                type: "Deleted",
                accept: deletedAcceptFresh,
                reject: vi.fn(),
              },
            ]
          : [],
      load: vi.fn(),
    }));

    context.document.body.getTrackedChanges.mockImplementation(() => ({
      items: [],
      load: vi.fn(),
    }));

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder).toEqual([
      "accept-added-initial",
      "accept-deleted-fresh",
    ]);
  });

  it("returns error and skips cleanup when a replace still appears pending after execution", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-half-after-success",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let phase = 0;
    const deletedAcceptSpy = vi.fn(() => {
      phase = 1;
    });
    const addedAcceptSpy = vi.fn(() => {
      phase = 2;
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-half-after-success",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-half-after-success",
        insertedTag: "stylistic:track-change:chunk0-half-after-success",
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
                accept: addedAcceptSpy,
                reject: vi.fn(),
              },
              {
                id: "tc-deleted-initial",
                type: "Deleted",
                accept: deletedAcceptSpy,
                reject: vi.fn(),
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
              accept: addedAcceptSpy,
              reject: vi.fn(),
            },
          ],
          load: vi.fn(),
        };
      }

      if (phase >= 2) {
        return {
          items: [
            {
              id: "tc-added-still-pending",
              type: "Added",
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

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("error");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.error).toContain("falso success");
    expect(context._cc.delete).not.toHaveBeenCalled();
  });

  it("retries one atomic accept batch when a fresh post-execute observation still exposes the full replace pair", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-post-execute-atomic-accept",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let phase = 0;
    const queuedTypes: string[] = [];
    const callOrder: string[] = [];

    const addedAcceptInitial = vi.fn(() => {
      queuedTypes.push("Added");
      callOrder.push("accept-added-initial");
      phase = 1;
    });
    const deletedAcceptInitial = vi.fn(() => {
      queuedTypes.push("Deleted");
      callOrder.push("accept-deleted-initial");
      phase = 2;
    });
    const addedAcceptAtomic = vi.fn(() => {
      queuedTypes.push("Added");
      callOrder.push("accept-added-atomic");
    });
    const deletedAcceptAtomic = vi.fn(() => {
      queuedTypes.push("Deleted");
      callOrder.push("accept-deleted-atomic");
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-post-execute-atomic-accept",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-post-execute-atomic-accept",
        insertedTag: "stylistic:track-change:chunk0-post-execute-atomic-accept",
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
                accept: addedAcceptInitial,
                reject: vi.fn(),
              },
              {
                id: "tc-deleted-initial",
                type: "Deleted",
                accept: deletedAcceptInitial,
                reject: vi.fn(),
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
              accept: deletedAcceptInitial,
              reject: vi.fn(),
            },
          ],
          load: vi.fn(),
        };
      }

      if (phase === 2) {
        return {
          items: [
            {
              id: "tc-added-atomic",
              type: "Added",
              accept: addedAcceptAtomic,
              reject: vi.fn(),
            },
            {
              id: "tc-deleted-atomic",
              type: "Deleted",
              accept: deletedAcceptAtomic,
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

    context.sync.mockImplementation(async () => {
      const batchKey = queuedTypes.join(",");

      if (phase === 2 && batchKey === "Added,Deleted") {
        queuedTypes.length = 0;
        phase = 3;
        return;
      }

      queuedTypes.length = 0;
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.error).toBeUndefined();
    expect(callOrder).toEqual([
      "accept-added-initial",
      "accept-deleted-initial",
      "accept-added-atomic",
      "accept-deleted-atomic",
    ]);
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("keeps accept replace successful when later stale atomic-retry setup is no longer needed after Added-first semantic resolution", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-post-execute-atomic-accept-fresh-pair",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let phase = 0;
    const queuedTypes: string[] = [];
    const callOrder: string[] = [];

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-post-execute-atomic-accept-fresh-pair",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-post-execute-atomic-accept-fresh-pair",
        insertedTag:
          "stylistic:track-change:chunk0-post-execute-atomic-accept-fresh-pair",
        deletedValue: "ni Shu",
        anchorValue: "No sabían si venía de ni Shu o de otro sitio.",
      }),
      deletedSideText: "ni Shu",
      comments: [],
    });

    const getCcRange = context._cc.getRange as unknown as () => {
      getTrackedChanges: ReturnType<typeof vi.fn>;
    };
    const ccRange = getCcRange();
    ccRange.getTrackedChanges.mockImplementation(() => {
      if (phase === 0) {
        return {
          items: [
            {
              id: "tc-added-initial",
              type: "Added",
              accept: vi.fn(() => {
                callOrder.push("accept-added-initial");
                phase = 1;
              }),
              reject: vi.fn(),
            },
            {
              id: "tc-deleted-initial",
              type: "Deleted",
              accept: vi.fn(() => {
                callOrder.push("accept-deleted-initial");
                phase = 2;
              }),
              reject: vi.fn(),
            },
          ],
          load: vi.fn(),
        };
      }

      if (phase === 2) {
        return {
          items: [
            {
              id: "tc-added-atomic-stale-range",
              type: "Added",
              accept: vi.fn(() => {
                queuedTypes.push("Added");
                callOrder.push("accept-added-atomic-stale");
              }),
              reject: vi.fn(),
            },
            {
              id: "tc-deleted-atomic-stale-range",
              type: "Deleted",
              accept: vi.fn(() => {
                queuedTypes.push("Deleted");
                callOrder.push("accept-deleted-atomic-stale");
              }),
              reject: vi.fn(),
            },
          ],
          load: vi.fn(),
        };
      }

      if (phase === 3) {
        return {
          items: [
            {
              id: "tc-deleted-fresh",
              type: "Deleted",
              accept: vi.fn(() => {
                queuedTypes.push("Deleted");
                callOrder.push("accept-deleted-stepwise-fresh");
              }),
              reject: vi.fn(),
            },
          ],
          load: vi.fn(),
        };
      }

      if (phase === 4) {
        return {
          items: [
            {
              id: "tc-added-fresh",
              type: "Added",
              accept: vi.fn(() => {
                queuedTypes.push("Added");
                callOrder.push("accept-added-stepwise-fresh");
              }),
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

    context.sync.mockImplementation(async () => {
      const batchKey = queuedTypes.join(",");

      if (phase === 2 && batchKey === "Added,Deleted") {
        queuedTypes.length = 0;
        phase = 3;
        throw new Error("ItemNotFound");
      }

      if (phase === 3 && batchKey === "Deleted") {
        queuedTypes.length = 0;
        phase = 4;
        return;
      }

      if (phase === 4 && batchKey === "Added") {
        queuedTypes.length = 0;
        phase = 5;
        return;
      }

      queuedTypes.length = 0;
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.error).toBeUndefined();
    expect(callOrder).toEqual(["accept-added-initial"]);
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("keeps accepted semantic resolution when telemetry emission fails", async () => {
    const suggestion = makeSuggestion({
      id: "s-telemetry-1",
      anchor: "texto original",
      suggestedText: "texto sugerido",
      context: "Contexto con texto original.",
    });
    const telemetryPort: ITelemetryPort = {
      emit: vi.fn().mockRejectedValue(new Error("telemetry sink offline")),
    };
    adapter = new WordAdapter(undefined, telemetryPort);

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-telemetry-1",
        insertedTag: "stylistic:track-change:s-telemetry-1",
        deletedValue: "texto original",
        anchorValue: "Contexto con texto original.",
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

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(result.commentDeleted).toBe(false);
    expect(result.error).toBeUndefined();
    expect(telemetryPort.emit).toHaveBeenCalled();
  });

  it("emits observation telemetry with replace debug metadata before execution", async () => {
    const suggestion = makeSuggestion({
      id: "s-observation-telemetry",
      anchor: "desde allí",
      suggestedText: "desde entonces",
      context:
        "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
    });
    const telemetryPort: ITelemetryPort = {
      emit: vi.fn().mockResolvedValue(undefined),
    };
    adapter = new WordAdapter(undefined, telemetryPort);

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-observation-telemetry",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-observation-telemetry",
        insertedTag: "stylistic:track-change:s-observation-telemetry",
        deletedValue: "desde allí",
        anchorValue:
          "Después de eso desperté en unas instalaciones de WEPO lejos de las garras de Jack y desde allí no volví a saber de él.",
      }),
      spanTCItems: [
        {
          id: "tc-added",
          type: "Added",
          accept: vi.fn(),
          reject: vi.fn(),
        },
      ],
      deletedSideText: "desde allí",
      deletedSideRangeTCItems: [
        {
          id: "tc-deleted",
          type: "Deleted",
          accept: vi.fn(),
          reject: vi.fn(),
        },
      ],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(telemetryPort.emit).toHaveBeenNthCalledWith(
      4,
      expect.objectContaining({
        phase: "observe-before",
        outcome: "succeeded",
        metadata: expect.objectContaining({
          trackedChangesObserved: 2,
          trackedChangeTypes: "Added,Deleted",
          selectedDeletedSource: "deletedSide",
          selectedAddedSource: "cc",
          selectedCcTag: "stylistic:track-change:s-observation-telemetry",
          selectedCcTitleKind: "compound-v2",
          deletedSideTrackedChangesCount: 1,
          deletedSideLocatorFound: true,
        }),
      }),
    );
  });
});
