import { describe, expect, it, vi } from "vitest";
import { DocumentReviewStateInspector } from "./DocumentReviewStateInspector";

describe("DocumentReviewStateInspector", () => {
  it("builds the safest empty fallback state", () => {
    const inspector = new DocumentReviewStateInspector();

    expect(inspector.buildEmptyState()).toEqual({
      pendingStylisticArtifacts: 0,
      hasPendingStylisticArtifacts: false,
      trackChangesActive: false,
    });
  });

  it("falls back to pendingBefore when reject post-resolution inspection throws", async () => {
    const inspector = new DocumentReviewStateInspector();
    const pendingBefore = {
      pendingStylisticArtifacts: 2,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    };
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

    const result = await inspector.inspectAfterResolution(
      context,
      pendingBefore,
      "reject",
      "s-1",
    );

    expect(result).toEqual({
      pendingAfter: pendingBefore,
      warning: {
        code: "inspection-failed",
        phase: "inspect-after",
        message: "GeneralException",
      },
    });
  });
});
