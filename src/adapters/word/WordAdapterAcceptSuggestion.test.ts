import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WordAdapter } from "./WordAdapter";
import {
  installWordWithContext,
  makeCompoundV2Title,
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

  it("fails fast when a track-change suggestion violates the replace contract", async () => {
    const suggestion = makeSuggestion({
      id: "s-invalid-track-change-contract",
      anchor: "texto original",
      suggestedText: "",
      context: "Contexto con texto original.",
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:s-invalid-track-change-contract",
      spanTCItems: [],
      comments: [],
    });

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("error");
    expect(result.error).toContain("Contrato invalido de track-change");
    expect(context._cc.getTrackedChanges).not.toHaveBeenCalled();
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
    expect(result.trackedChangesAffected).toBe(1);
    expect(result.error).toContain("falso success");
    expect(context._cc.delete).not.toHaveBeenCalled();
  });

  it("fails closed when a fresh post-execute observation still exposes the full replace pair", async () => {
    const suggestion = makeSuggestion({
      id: "chunk0-post-execute-full-pair-still-pending-accept",
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
    const deletedAcceptInitial = vi.fn(() => {
      callOrder.push("accept-deleted-initial");
      phase = 2;
    });

    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: "stylistic:track-change:chunk0-post-execute-full-pair-still-pending-accept",
      ccTitle: makeCompoundV2Title({
        suggestionId: "chunk0-post-execute-full-pair-still-pending-accept",
        insertedTag:
          "stylistic:track-change:chunk0-post-execute-full-pair-still-pending-accept",
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

      if (phase >= 2) {
        return {
          items: [
            {
              id: "tc-added-still-pending",
              type: "Added",
              accept: vi.fn(() => {
                callOrder.push("unexpected-accept-added-extra-attempt");
              }),
              reject: vi.fn(),
            },
            {
              id: "tc-deleted-still-pending",
              type: "Deleted",
              accept: vi.fn(() => {
                callOrder.push("unexpected-accept-deleted-extra-attempt");
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

    installWordWithContext(context);

    const result = await adapter.acceptSuggestion(suggestion);

    expect(result.status).toBe("error");
    expect(result.error).toBeTruthy();
    expect(result.commentDeleted).toBe(false);
    expect(context._cc.delete).not.toHaveBeenCalled();
    expect(callOrder).toEqual([
      "accept-added-initial",
      "accept-deleted-initial",
    ]);
  });
});
