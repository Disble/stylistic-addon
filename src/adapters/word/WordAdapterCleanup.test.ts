import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getCleanupMocks } from "./WordAdapterTestHelper";
import { WordAdapter } from "./WordAdapter";

describe("WordAdapter.cleanupResolvedComments", () => {
  const cleanupMocks = getCleanupMocks();
  let adapter: WordAdapter;

  beforeEach(() => {
    vi.clearAllMocks();
    adapter = new WordAdapter();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("delegates to CommentCleanup and returns its counts", async () => {
    cleanupMocks.cleanupResolvedComments.mockResolvedValueOnce({
      deleted: 3,
      kept: 1,
    });

    await expect(adapter.cleanupResolvedComments()).resolves.toEqual({
      deleted: 3,
      kept: 1,
    });

    expect(cleanupMocks.cleanupResolvedComments).toHaveBeenCalledOnce();
  });

  it("propagates cleanup errors", async () => {
    cleanupMocks.cleanupResolvedComments.mockRejectedValueOnce(
      new Error("cleanup failed"),
    );

    await expect(adapter.cleanupResolvedComments()).rejects.toThrow(
      "cleanup failed",
    );
  });
});
