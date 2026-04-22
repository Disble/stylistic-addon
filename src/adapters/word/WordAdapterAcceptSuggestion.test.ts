import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "./WordAdapter";
import {
  installRejectingWord,
  makeCompoundV2Title,
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

  it("returns identity-lost when compound-v2 metadata is corrupt or incomplete", async () => {
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
    expect(result.pendingAfter.hasPendingStylisticArtifacts).toBe(true);
    expect(result.error).toBeUndefined();
    expect(tcAcceptSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });

  it("accepts and deletes the matching comment by content when authorName is not Stylistic", async () => {
    const suggestion = makeSuggestion();
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

  it("returns unobservable when compound-v2 only stores the original pre-replace context as operational anchor", async () => {
    const suggestion = makeSuggestion({
      id: "s-pre-replace-context",
      anchor: "parecía ser",
      suggestedText: "parece que es",
      context: "Hasta donde sabía, parecía ser la única al tanto.",
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-pre-replace-context",
      ccTitle: makeCompoundV2Title({
        suggestionId: "s-pre-replace-context",
        insertedTag: "stylistic:track-change:s-pre-replace-context",
        deletedValue: "parecía ser",
        anchorValue: "Hasta donde sabía, parecía ser la única al tanto.",
      }),
      spanTCItems: [],
      rangeTCItems: [],
      bodyTCItems: [],
      operationalAnchorText: undefined,
      operationalAnchorRangeTCItems: [],
      comments: [],
    });
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("unobservable");
    expect(result.error).toContain("Word no expuso suficientes tracked changes");
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
    const suggestion = makeSuggestion({ id: "s-final" });
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-final",
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

  it("returns accepted with warnings when a second accept click completes semantic resolution before late ItemNotFound cleanup failure", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-4",
      anchor: "sándwich o sánduche",
      suggestedText: "sándwich o sánguche",
      context: "Quería debatir si debía decir sándwich o sánduche antes de hacer el pedido.",
    });

    let resolutionPhase = 0;
    let cleanupAttemptedAfterFullResolution = false;
    let shouldThrowOnDeletedAccept = true;

    const addedAcceptSpy = vi.fn(() => {
      resolutionPhase = Math.max(resolutionPhase, 1);
    });
    const deletedAcceptSpy = vi.fn(() => {
      if (shouldThrowOnDeletedAccept) {
        shouldThrowOnDeletedAccept = false;
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
            accept: addedAcceptSpy,
            reject: vi.fn(),
          },
          {
            id: "tc-deleted",
            type: "Deleted",
            accept: deletedAcceptSpy,
            reject: vi.fn(),
          },
        ];
      }

      if (resolutionPhase === 1) {
        return [
          {
            id: "tc-deleted",
            type: "Deleted",
            accept: deletedAcceptSpy,
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

    const firstResult = await adapter.acceptSuggestion(suggestion);
    const secondResult = await adapter.acceptSuggestion(suggestion);

    expect(firstResult.status).toBe("error");
    expect(firstResult.error).toContain("ItemNotFound");
    expect(secondResult.status).toBe("accepted");
    expect(secondResult.trackedChangesAffected).toBe(1);
    expect(secondResult.commentDeleted).toBe(false);
    expect(secondResult.executionReport).toEqual({
      attempted: 1,
      completed: 1,
      remaining: 0,
    });
    expect(secondResult.warnings).toHaveLength(3);
    expect(secondResult.warnings).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: "cleanup-failed",
          phase: "cleanup",
        }),
        expect.objectContaining({
          code: "inspection-failed",
          phase: "inspect-after",
        }),
      ]),
    );
    expect(addedAcceptSpy).toHaveBeenCalledOnce();
    expect(deletedAcceptSpy).toHaveBeenCalledTimes(2);
    expect(commentDeleteSpy).toHaveBeenCalledOnce();
    expect(context._cc.delete).toHaveBeenCalledWith(true);
  });
});
