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
  content?: string;
  delete: ReturnType<typeof vi.fn>;
  getRange: ReturnType<typeof vi.fn>;
};

type CleanupContentControl = {
  tag: string;
  getRange: ReturnType<typeof vi.fn>;
};

function makeComment(
  authorName: string,
  range?: CleanupRange,
  content?: string,
): CleanupComment {
  const actualRange =
    range ??
    ({
      compareLocationWith: vi.fn(),
    } satisfies CleanupRange);

  return {
    authorName,
    content,
    delete: vi.fn(),
    getRange: vi.fn(() => actualRange),
  };
}

function makeCC(tag: string, range?: CleanupRange): CleanupContentControl {
  const actualRange =
    range ??
    ({
      compareLocationWith: vi.fn(),
    } satisfies CleanupRange);

  return {
    tag,
    getRange: vi.fn(() => actualRange),
  };
}

function installWordCleanupContext(
  options: {
    comments?: CleanupComment[];
    contentControls?: CleanupContentControl[];
  } = {},
) {
  const commentsCollection = {
    items: options.comments ?? [],
    load: vi.fn(),
  };

  const contentControlsCollection = {
    items: options.contentControls ?? [],
    load: vi.fn(),
  };

  const context = {
    document: {
      body: {
        getComments: vi.fn(() => commentsCollection),
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

  // CC-03 — Empty document (no comments)
  it("returns { deleted: 0, kept: 0 } when there are no comments", async () => {
    const { context } = installWordCleanupContext({
      comments: [],
      contentControls: [],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 0, kept: 0 });
    expect(context.sync).toHaveBeenCalledTimes(1);
  });

  // CC-03 variant — non-Stylistic comments are ignored
  it("returns zero counts when there are no Stylistic comments", async () => {
    const otherComment = makeComment("Someone Else");
    const { context, commentsCollection } = installWordCleanupContext({
      comments: [otherComment],
      contentControls: [],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 0, kept: 0 });
    expect(commentsCollection.load).toHaveBeenCalledWith({
      select: "authorName,content",
    });
    expect(otherComment.delete).not.toHaveBeenCalled();
    expect(context.sync).toHaveBeenCalledTimes(1);
  });

  // CC-02 — Track-change comment with no colocated CC is deleted
  it("deletes every Stylistic comment when no active Stylistic CCs exist", async () => {
    const stylisticA = makeComment("Usuario de prueba", undefined, "[Claridad]\nMas claro");
    const stylisticB = makeComment("Usuario de prueba", undefined, "[Registro]\nMas natural");
    const foreignComment = makeComment("Reviewer");
    // CC with foreign tag does NOT protect Stylistic comments
    const foreignCC = makeCC("other-tool:chunk0-0");
    const { context } = installWordCleanupContext({
      comments: [stylisticA, stylisticB, foreignComment],
      contentControls: [foreignCC],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 2, kept: 0 });
    expect(stylisticA.delete).toHaveBeenCalledOnce();
    expect(stylisticB.delete).toHaveBeenCalledOnce();
    expect(foreignComment.delete).not.toHaveBeenCalled();
    // Sync 1: load; Sync 2: getRange; Sync 3: delete comments
    expect(context.sync).toHaveBeenCalledTimes(3);
  });

  // CC-01 — Comment colocated with active Stylistic CC is kept
  it("keeps comments whose ranges overlap a stylistic track-change CC", async () => {
    const ccRange = {
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

    const keptComment = makeComment(
      "Usuario de prueba",
      overlappingCommentRange,
      "[Claridad]\nMas claro",
    );
    const deletedComment = makeComment(
      "Usuario de prueba",
      distantCommentRange,
      "[Registro]\nMas natural",
    );
    const commentOnlyCC = makeCC("stylistic:track-change:chunk0-0", ccRange);
    const { context } = installWordCleanupContext({
      comments: [keptComment, deletedComment],
      contentControls: [commentOnlyCC],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 1, kept: 1 });
    expect(keptComment.getRange).toHaveBeenCalled();
    expect(deletedComment.getRange).toHaveBeenCalled();
    expect(commentOnlyCC.getRange).toHaveBeenCalledOnce();
    expect(overlappingCommentRange.compareLocationWith).toHaveBeenCalledWith(
      ccRange,
    );
    expect(distantCommentRange.compareLocationWith).toHaveBeenCalledWith(
      ccRange,
    );
    expect(keptComment.delete).not.toHaveBeenCalled();
    expect(deletedComment.delete).toHaveBeenCalledOnce();
    // Sync 1: load; Sync 2: getRange for comments + CCs; Sync 3: spatial comparisons; Sync 4: deletes
    expect(context.sync).toHaveBeenCalledTimes(4);
  });

  it("deletes Stylistic-shaped comments even when Word reports the current user as author", async () => {
    const stylisticComment = makeComment(
      "Usuario de prueba",
      undefined,
      "[gramática]\nRedundancia pronominal.",
    );
    installWordCleanupContext({
      comments: [stylisticComment],
      contentControls: [],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 1, kept: 0 });
    expect(stylisticComment.delete).toHaveBeenCalledOnce();
  });

  it("treats every relation outside the overlap allowlist as resolved", async () => {
    const ccRange = {
      compareLocationWith: vi.fn(),
    } satisfies CleanupRange;
    const edgeComment = makeComment(
      "Usuario de prueba",
      {
        compareLocationWith: vi.fn(
          () => ({ value: "AdjacentBefore" }) satisfies FakeClientResult,
        ),
      },
      "[Claridad]\nMas claro",
    );
    installWordCleanupContext({
      comments: [edgeComment],
      contentControls: [makeCC("stylistic:track-change:chunk0-1", ccRange)],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 1, kept: 0 });
    expect(edgeComment.delete).toHaveBeenCalledOnce();
  });

  it("propagates Word.run errors instead of swallowing them", async () => {
    const failure = new Error("Word sync failed");
    const comment = makeComment("Usuario de prueba", undefined, "[Claridad]\nMas claro");
    const { context } = installWordCleanupContext({
      comments: [comment],
      contentControls: [],
    });
    context.sync.mockRejectedValueOnce(failure);

    await expect(cleanupResolvedComments()).rejects.toThrow("Word sync failed");
  });

  // CC-01 extended — multiple CCs, comment overlaps one of them
  it("keeps a comment when it overlaps any of multiple active Stylistic CCs", async () => {
    const ccRange1 = { compareLocationWith: vi.fn() } satisfies CleanupRange;
    const ccRange2 = { compareLocationWith: vi.fn() } satisfies CleanupRange;
    const commentRange = {
      compareLocationWith: vi
        .fn()
        .mockReturnValueOnce({ value: "Before" } satisfies FakeClientResult) // vs CC1
        .mockReturnValueOnce({ value: "Equal" } satisfies FakeClientResult), // vs CC2
    } satisfies CleanupRange;

    const keptComment = makeComment(
      "Usuario de prueba",
      commentRange,
      "[Claridad]\nMas claro",
    );
    installWordCleanupContext({
      comments: [keptComment],
      contentControls: [
        makeCC("stylistic:comment-only:chunk0-0", ccRange1),
        makeCC("stylistic:track-change:chunk0-1", ccRange2),
      ],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 0, kept: 1 });
    expect(keptComment.delete).not.toHaveBeenCalled();
  });

  // CC-02 — All comments deleted when there are no active Stylistic CCs at all
  it("deletes all Stylistic comments when contentControls collection is empty", async () => {
    const commentA = makeComment("Usuario de prueba", undefined, "[Claridad]\nMas claro");
    const commentB = makeComment("Usuario de prueba", undefined, "[Registro]\nMas natural");
    installWordCleanupContext({
      comments: [commentA, commentB],
      contentControls: [],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 2, kept: 0 });
    expect(commentA.delete).toHaveBeenCalledOnce();
    expect(commentB.delete).toHaveBeenCalledOnce();
  });

  it("keeps comments attached to unresolved track-change suggestions and deletes only true orphans", async () => {
    const activeRange = {
      compareLocationWith: vi.fn(
        () => ({ value: "Equal" }) satisfies FakeClientResult,
      ),
    } satisfies CleanupRange;
    const orphanRange = {
      compareLocationWith: vi.fn(
        () => ({ value: "Before" }) satisfies FakeClientResult,
      ),
    } satisfies CleanupRange;

    const activeComment = makeComment(
      "Usuario de prueba",
      activeRange,
      "[Gramática]\nComentario activo",
    );
    const orphanComment = makeComment(
      "Usuario de prueba",
      orphanRange,
      "[Estilo]\nComentario huérfano",
    );

    installWordCleanupContext({
      comments: [activeComment, orphanComment],
      contentControls: [makeCC("stylistic:track-change:s-1")],
    });

    const result = await cleanupResolvedComments();

    expect(result).toEqual({ deleted: 1, kept: 1 });
    expect(activeComment.delete).not.toHaveBeenCalled();
    expect(orphanComment.delete).toHaveBeenCalledOnce();
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
