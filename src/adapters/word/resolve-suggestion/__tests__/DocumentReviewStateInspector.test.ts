import { describe, expect, it, vi } from "vitest";
import { DocumentReviewStateInspector } from "../DocumentReviewStateInspector";

describe("DocumentReviewStateInspector", () => {
  it("builds the safest empty fallback state", () => {
    const inspector = new DocumentReviewStateInspector();

    expect(inspector.buildEmptyState()).toEqual({
      pendingStylisticArtifacts: 0,
      hasPendingStylisticArtifacts: false,
      trackChangesActive: false,
    });
  });

  it("propagates reject post-resolution inspection failures", async () => {
    const inspector = new DocumentReviewStateInspector();
    const context = {
      document: {
        contentControls: {
          load: vi.fn(() => {
            throw new Error("GeneralException");
          }),
        },
        load: vi.fn(),
      },
      sync: vi.fn(),
    } as unknown as Word.RequestContext;

    await expect(inspector.inspectAfterResolution(context)).rejects.toThrow("GeneralException");
  });
});
