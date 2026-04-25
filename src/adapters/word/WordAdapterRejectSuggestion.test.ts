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

  it("fails closed without retrying a fresh Deleted candidate after the first Deleted proxy fails", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-no-recovery-reject",
      anchor: "ni Shu",
      suggestedText: "ni de Shu",
      context: "No sabían si venía de ni Shu o de otro sitio.",
    });

    let observationCount = 0;
    const deletedRejectStale = vi.fn(() => {
      throw new Error("ItemNotFound");
    });
    const deletedRejectFresh = vi.fn();

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-no-recovery-reject",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-no-recovery-reject",
        insertedTag: "stylistic:track-change:chunk0-no-recovery-reject",
        deletedValue: "ni Shu",
        anchorValue: "No sabían si venía de ni Shu o de otro sitio.",
      }),
      comments: [],
    });

    context._cc.getTrackedChanges.mockImplementation(() => {
      observationCount += 1;
      return {
        items: [
          {
            id: observationCount === 1 ? "tc-deleted-stale" : "tc-deleted-fresh",
            type: "Deleted",
            accept: vi.fn(),
            reject:
              observationCount === 1 ? deletedRejectStale : deletedRejectFresh,
          },
          {
            id: "tc-added-pending",
            type: "Added",
            accept: vi.fn(),
            reject: vi.fn(),
          },
        ],
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

    expect(result.status).toBe("error");
    expect(deletedRejectStale).toHaveBeenCalledOnce();
    expect(deletedRejectFresh).not.toHaveBeenCalled();
    expect(result.error).toBeTruthy();
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
    expect(result.trackedChangesAffected).toBe(1);
    expect(result.error).toContain("falso success");
    expect(context._cc.delete).not.toHaveBeenCalled();
  });
});
