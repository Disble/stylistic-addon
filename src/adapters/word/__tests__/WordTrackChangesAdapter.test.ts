import { afterEach, describe, expect, it, vi } from "vitest";
import { installWordWithContext } from "./WordAdapterTestHelper";
import { WordTrackChangesAdapter } from "../WordTrackChangesAdapter";

/** Builds a minimal request context mock accepted by track-changes adapter tests. */
function makeTrackChangesContext(changeTrackingMode: string): Word.RequestContext {
  return {
    document: {
      contentControls: { items: [], load: vi.fn() },
      load: vi.fn(),
      changeTrackingMode,
    },
    sync: vi.fn().mockResolvedValue(undefined),
  } as unknown as Word.RequestContext;
}

describe("WordTrackChangesAdapter", () => {
  const adapter = new WordTrackChangesAdapter();

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("derives the current document review state from stylistic content controls", async () => {
    const context = {
      document: {
        contentControls: {
          items: [{ tag: "stylistic:track-change:s1" }, { tag: "other" }],
          load: vi.fn(),
        },
        load: vi.fn(),
        changeTrackingMode: "trackAll",
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };

    installWordWithContext(context);

    await expect(adapter.getDocumentReviewState()).resolves.toEqual({
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    });
  });

  it("enables track changes only when currently off", async () => {
    const context = makeTrackChangesContext("off");

    await expect(adapter.ensureTrackChangesActive(context)).resolves.toBe(true);
    expect(context.document.changeTrackingMode).toBe("trackAll");
  });
});
