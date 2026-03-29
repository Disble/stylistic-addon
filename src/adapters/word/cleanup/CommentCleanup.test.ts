import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  cleanupResolvedComments,
  OVERLAPPING_RELATIONS,
} from "./CommentCleanup";

type FakeClientResult = { value: string };

type CleanupRange = {
  compareLocationWith: ReturnType<typeof vi.fn>;
};

type CleanupComment = {
  authorName: string;
  delete: ReturnType<typeof vi.fn>;
  getRange: ReturnType<typeof vi.fn>;
};

type CleanupTrackedChange = {
  author: string;
  type?: string;
  getRange: ReturnType<typeof vi.fn>;
};

function makeComment(authorName: string, range?: CleanupRange): CleanupComment {
  const actualRange =
    range ??
    ({
      compareLocationWith: vi.fn(),
    } satisfies CleanupRange);

  return {
    authorName,
    delete: vi.fn(),
    getRange: vi.fn(() => actualRange),
  };
}

function makeTrackedChange(
  author: string,
  range?: CleanupRange,
): CleanupTrackedChange {
  const actualRange =
    range ??
    ({
      compareLocationWith: vi.fn(),
    } satisfies CleanupRange);

  return {
    author,
    type: "Replace",
    getRange: vi.fn(() => actualRange),
  };
}

function installWordCleanupContext(
  options: {
    comments?: CleanupComment[];
    trackedChanges?: CleanupTrackedChange[];
    contentControlTags?: string[];
  } = {},
) {
  const commentsCollection = {
    items: options.comments ?? [],
    load: vi.fn(),
  };

  const trackedCollection = {
    items: options.trackedChanges ?? [],
    load: vi.fn(),
  };

  const contentControlsCollection = {
    items: (options.contentControlTags ?? []).map((tag) => ({
      tag,
      getRange: vi.fn(() => ({ compareLocationWith: vi.fn() })),
    })),
    load: vi.fn(),
  };

  const context = {
    document: {
      body: {
        getComments: vi.fn(() => commentsCollection),
        getTrackedChanges: vi.fn(() => trackedCollection),
      },
      contentControls: contentControlsCollection,
    },
    sync: vi.fn(async () => undefined),
  };

  (globalThis as any).Word = {
    run: vi.fn(async (callback: (ctx: typeof context) => unknown) =>
      callback(context),
    ),
  };

  return {
    context,
    commentsCollection,
    trackedCollection,
    contentControlsCollection,
  };
}

describe("cleanupResolvedComments", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    delete (globalThis as any).Word;
  });

  it("returns zero counts when there are no Stylistic comments", async () => {
    const otherComment = makeComment("Someone Else");
    const otherTc = makeTrackedChange("Stylistic");
    const { context, commentsCollection, trackedCollection } =
      installWordCleanupContext({
        comments: [otherComment],
        trackedChanges: [otherTc],
      });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 0, kept: 0 });
    expect(trackedCollection.load).toHaveBeenCalledWith({
      select: "author,type",
    });
    expect(commentsCollection.load).toHaveBeenCalledWith({
      select: "authorName",
    });
    expect(otherComment.delete).not.toHaveBeenCalled();
    expect(context.sync).toHaveBeenCalledTimes(1);
  });

  it("deletes every Stylistic comment when no Stylistic tracked changes remain", async () => {
    const stylisticA = makeComment("Stylistic");
    const stylisticB = makeComment("Stylistic");
    const foreignComment = makeComment("Reviewer");
    const foreignTc = makeTrackedChange("Reviewer");
    const { context } = installWordCleanupContext({
      comments: [stylisticA, stylisticB, foreignComment],
      trackedChanges: [foreignTc],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 2, kept: 0 });
    expect(stylisticA.delete).toHaveBeenCalledOnce();
    expect(stylisticB.delete).toHaveBeenCalledOnce();
    expect(foreignComment.delete).not.toHaveBeenCalled();
    // Sync 1: load collections; Sync 2: getRange for CC-check; Sync 3: delete comments
    expect(context.sync).toHaveBeenCalledTimes(3);
  });

  it("keeps comments whose ranges overlap a Stylistic tracked change", async () => {
    const tcRange = {
      compareLocationWith: vi.fn(),
    } satisfies CleanupRange;
    const overlappingCommentRange = {
      compareLocationWith: vi.fn(
        () => ({ value: "Contains" }) satisfies FakeClientResult,
      ),
    } satisfies CleanupRange;
    const distantCommentRange = {
      compareLocationWith: vi.fn(
        () => ({ value: "Before" }) satisfies FakeClientResult,
      ),
    } satisfies CleanupRange;

    const keptComment = makeComment("Stylistic", overlappingCommentRange);
    const deletedComment = makeComment("Stylistic", distantCommentRange);
    const trackedChange = makeTrackedChange("Stylistic", tcRange);
    const { context } = installWordCleanupContext({
      comments: [keptComment, deletedComment],
      trackedChanges: [trackedChange],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 1, kept: 1 });
    // getRange called twice per comment: once for CC-check (Sync 2) and once for TC colocation (Sync 4)
    expect(keptComment.getRange).toHaveBeenCalledTimes(2);
    expect(deletedComment.getRange).toHaveBeenCalledTimes(2);
    expect(trackedChange.getRange).toHaveBeenCalledOnce();
    expect(overlappingCommentRange.compareLocationWith).toHaveBeenCalledWith(
      tcRange,
    );
    expect(distantCommentRange.compareLocationWith).toHaveBeenCalledWith(
      tcRange,
    );
    expect(keptComment.delete).not.toHaveBeenCalled();
    expect(deletedComment.delete).toHaveBeenCalledOnce();
    // Sync 1: load; Sync 2: CC-check getRange (no CCs → skip comparison sync);
    // Sync 3: TC comment ranges + TC ranges; Sync 4: spatial comparisons; Sync 5: deletes
    expect(context.sync).toHaveBeenCalledTimes(5);
  });

  it("treats every relation outside the overlap allowlist as resolved", async () => {
    const tcRange = {
      compareLocationWith: vi.fn(),
    } satisfies CleanupRange;
    const edgeComment = makeComment("Stylistic", {
      compareLocationWith: vi.fn(
        () => ({ value: "AdjacentBefore" }) satisfies FakeClientResult,
      ),
    });
    installWordCleanupContext({
      comments: [edgeComment],
      trackedChanges: [makeTrackedChange("Stylistic", tcRange)],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 1, kept: 0 });
    expect(edgeComment.delete).toHaveBeenCalledOnce();
  });

  it("propagates Word.run errors instead of swallowing them", async () => {
    const failure = new Error("Word sync failed");
    const comment = makeComment("Stylistic");
    const trackedChange = makeTrackedChange("Stylistic");
    const { context } = installWordCleanupContext({
      comments: [comment],
      trackedChanges: [trackedChange],
    });
    context.sync.mockRejectedValueOnce(failure);

    await expect(cleanupResolvedComments()).rejects.toThrow("Word sync failed");
  });

  it("does not delete comments owned by active comment-only CCs even when there are zero TCs", async () => {
    // A Stylistic comment whose range overlaps the comment-only CC
    const commentOnlyCommentRange = {
      compareLocationWith: vi.fn(
        () => ({ value: "Equal" }) satisfies FakeClientResult,
      ),
    } satisfies CleanupRange;
    const commentOnlyComment = makeComment(
      "Stylistic",
      commentOnlyCommentRange,
    );

    const { context } = installWordCleanupContext({
      comments: [commentOnlyComment],
      trackedChanges: [], // zero TCs → would normally short-circuit and delete all
      contentControlTags: ["stylistic:comment-only:chunk0-0"],
    });

    const result = await cleanupResolvedComments();

    // The comment is protected — nothing deleted, one kept
    expect(result).toEqual({ deleted: 0, kept: 1 });
    expect(commentOnlyComment.delete).not.toHaveBeenCalled();
    // Sync must not call context.sync for short-circuit TC delete (that branch is skipped)
    expect(context.sync).toHaveBeenCalled();
  });

  it("deletes only orphaned track-change comments when all TCs are resolved but comment-only CCs remain", async () => {
    // comment-only CC still active — its comment must be preserved
    const commentOnlyCommentRange = {
      compareLocationWith: vi.fn(
        () => ({ value: "Equal" }) satisfies FakeClientResult,
      ),
    } satisfies CleanupRange;
    const commentOnlyComment = makeComment(
      "Stylistic",
      commentOnlyCommentRange,
    );

    // track-change comment that has no remaining TC
    const orphanedTcCommentRange = {
      compareLocationWith: vi.fn(
        () => ({ value: "Before" }) satisfies FakeClientResult,
      ),
    } satisfies CleanupRange;
    const orphanedTcComment = makeComment("Stylistic", orphanedTcCommentRange);

    installWordCleanupContext({
      comments: [commentOnlyComment, orphanedTcComment],
      trackedChanges: [], // zero TCs → track-change comments are orphans
      contentControlTags: ["stylistic:comment-only:chunk0-0"],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 1, kept: 1 });
    expect(commentOnlyComment.delete).not.toHaveBeenCalled();
    expect(orphanedTcComment.delete).toHaveBeenCalledOnce();
  });
});

describe("OVERLAPPING_RELATIONS", () => {
  it("should be exported and contain expected relation types", () => {
    expect(OVERLAPPING_RELATIONS).toBeDefined();
    expect(Array.isArray(OVERLAPPING_RELATIONS)).toBe(true);
    expect(OVERLAPPING_RELATIONS).toContain("Equal");
    expect(OVERLAPPING_RELATIONS).toContain("Contains");
    expect(OVERLAPPING_RELATIONS).toContain("Inside");
    expect(OVERLAPPING_RELATIONS).toContain("OverlapsBefore");
    expect(OVERLAPPING_RELATIONS).toContain("OverlapsAfter");
  });
});
