import { afterEach, describe, expect, it, vi } from "vitest";
import { installWordWithContext } from "../../__tests__/WordAdapterTestHelper";
import { DocumentReviewStateInspector } from "../DocumentReviewStateInspector";

describe("DocumentReviewStateInspector", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("ignores template content controls whose tag is null", async () => {
    const inspector = new DocumentReviewStateInspector();
    const context = {
      document: {
        contentControls: {
          items: [{ tag: null }, { tag: "stylistic:track-change:s1" }],
          load: vi.fn(),
        },
        load: vi.fn(),
        changeTrackingMode: "off",
      },
      sync: vi.fn().mockResolvedValue(undefined),
    };
    installWordWithContext(context);

    await expect(inspector.inspect(context as unknown as Word.RequestContext)).resolves.toEqual({
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: false,
    });
  });

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
