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
