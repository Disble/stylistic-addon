import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "./WordAdapter";
import {
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
    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("accepted");
    expect(result.trackedChangesAffected).toBe(2);
    expect(callOrder).toEqual(["accept-added-body", "accept-deleted-range"]);
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

});
