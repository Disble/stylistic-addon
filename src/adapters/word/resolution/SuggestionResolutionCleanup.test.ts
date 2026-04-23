import { describe, expect, it, vi } from "vitest";
import { SuggestionResolutionCleanup } from "./SuggestionResolutionCleanup";
import type { ColocatedCommentContext } from "./ResolutionContext";

/**
 * Builds a minimal Word.RequestContext stub for cleanup tests. The sync
 * function can be configured to throw on the next invocation to simulate
 * the host raising GeneralException after a previously successful mutation.
 */
function buildContextStub(): {
  context: Word.RequestContext;
  syncMock: ReturnType<typeof vi.fn>;
} {
  const syncMock = vi.fn(async () => undefined);
  const context = { sync: syncMock } as unknown as Word.RequestContext;
  return { context, syncMock };
}

function makeColocatedComment(
  deleteImpl: () => void = () => undefined,
): ColocatedCommentContext {
  return {
    comment: { delete: vi.fn(deleteImpl) } as unknown as Word.Comment,
    range: {} as unknown as Word.Range,
  };
}

describe("SuggestionResolutionCleanup soft-success contract", () => {
  it("returns commentDeleted=true when comment.delete()/sync raises GeneralException after the host already invalidated the comment proxy", async () => {
    const cleanup = new SuggestionResolutionCleanup("s-1", "reject");
    const { context, syncMock } = buildContextStub();
    syncMock.mockImplementationOnce(async () => {
      throw new Error("GeneralException");
    });
    const colocatedComment = makeColocatedComment();

    const commentDeleted = await cleanup.deleteLocatedStylisticComment(
      context,
      colocatedComment,
    );

    expect(commentDeleted).toBe(true);
  });

  it("propagates non-GeneralException errors from comment.delete()/sync so legitimate failures still abort the workflow", async () => {
    const cleanup = new SuggestionResolutionCleanup("s-2", "reject");
    const { context, syncMock } = buildContextStub();
    syncMock.mockImplementationOnce(async () => {
      throw new Error("InvalidOperationInCellEdit");
    });
    const colocatedComment = makeColocatedComment();

    await expect(
      cleanup.deleteLocatedStylisticComment(context, colocatedComment),
    ).rejects.toThrow("InvalidOperationInCellEdit");
  });

  it("treats CC anchor cleanup GeneralException as soft success (CC was already collapsed by prior tracked-change resolution)", async () => {
    const cleanup = new SuggestionResolutionCleanup("s-3", "reject");
    const { context, syncMock } = buildContextStub();
    syncMock.mockImplementationOnce(async () => {
      throw new Error("GeneralException");
    });
    const cc = {
      tag: "stylistic:track-change:s-3",
      delete: vi.fn(),
    } as unknown as Word.ContentControl;

    await expect(
      cleanup.cleanupResolvedSuggestionAnchor(context, cc),
    ).resolves.toBeUndefined();
  });

  it("propagates non-GeneralException errors from cc.delete()/sync so legitimate cleanup failures still abort the workflow", async () => {
    const cleanup = new SuggestionResolutionCleanup("s-4", "accept");
    const { context, syncMock } = buildContextStub();
    syncMock.mockImplementationOnce(async () => {
      throw new Error("InvalidObjectPath");
    });
    const cc = {
      tag: "stylistic:track-change:s-4",
      delete: vi.fn(),
    } as unknown as Word.ContentControl;

    await expect(
      cleanup.cleanupResolvedSuggestionAnchor(context, cc),
    ).rejects.toThrow("InvalidObjectPath");
  });
});
