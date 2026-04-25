import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  cleanupResolvedComments,
  getCleanupPreview,
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

/** Creates a fake cleanup comment with optional range/content overrides. */
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

/** Creates a fake active content control with a stable range. */
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

/** Installs a minimal Word.run cleanup context for preview/delete tests. */
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

  return { context, commentsCollection, contentControlsCollection };
}

describe("CommentCleanup", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
    delete (globalThis as any).Word;
  });

  it("returns zero preview/delete counts when there are no Stylistic comments", async () => {
    const foreignComment = makeComment("Reviewer");
    const { commentsCollection, context } = installWordCleanupContext({
      comments: [foreignComment],
      contentControls: [],
    });

    await expect(getCleanupPreview()).resolves.toEqual({
      deletable: 0,
      kept: 0,
    });
    await expect(cleanupResolvedComments()).resolves.toEqual({
      deleted: 0,
      kept: 0,
    });

    expect(commentsCollection.load).toHaveBeenCalledWith({
      select: "authorName,content",
    });
    expect(foreignComment.delete).not.toHaveBeenCalled();
    expect(context.sync).toHaveBeenCalled();
  });

  it("deletes every Stylistic comment when no active Stylistic content controls exist", async () => {
    const stylisticA = makeComment(
      "Usuario de prueba",
      undefined,
      "[Claridad]\nMas claro",
    );
    const stylisticB = makeComment(
      "Usuario de prueba",
      undefined,
      "[Registro]\nMas natural",
    );

    await installWordCleanupContext({
      comments: [stylisticA, stylisticB],
      contentControls: [makeCC("other-tool:chunk0-0")],
    });

    await expect(getCleanupPreview()).resolves.toEqual({
      deletable: 2,
      kept: 0,
    });
    await expect(cleanupResolvedComments()).resolves.toEqual({
      deleted: 2,
      kept: 0,
    });

    expect(stylisticA.delete).toHaveBeenCalledOnce();
    expect(stylisticB.delete).toHaveBeenCalledOnce();
  });

  it("keeps comments that overlap any active Stylistic content control and deletes true orphans", async () => {
    const installMixedColocationScenario = () => {
      const ccRange1 = { compareLocationWith: vi.fn() } satisfies CleanupRange;
      const ccRange2 = { compareLocationWith: vi.fn() } satisfies CleanupRange;
      const keptRange = {
        compareLocationWith: vi
          .fn()
          .mockReturnValueOnce({ value: "Before" } satisfies FakeClientResult)
          .mockReturnValueOnce({ value: "Equal" } satisfies FakeClientResult),
      } satisfies CleanupRange;
      const orphanRange = {
        compareLocationWith: vi
          .fn()
          .mockReturnValueOnce({ value: "Before" } satisfies FakeClientResult)
          .mockReturnValueOnce({ value: "AdjacentBefore" } satisfies FakeClientResult),
      } satisfies CleanupRange;

      const keptComment = makeComment(
        "Usuario de prueba",
        keptRange,
        "[Claridad]\nComentario activo",
      );
      const orphanComment = makeComment(
        "Usuario de prueba",
        orphanRange,
        "[Estilo]\nComentario huérfano",
      );

      installWordCleanupContext({
        comments: [keptComment, orphanComment],
        contentControls: [
          makeCC("stylistic:comment-only:chunk0-0", ccRange1),
          makeCC("stylistic:track-change:chunk0-1", ccRange2),
        ],
      });

      return { keptComment, orphanComment };
    };

    installMixedColocationScenario();
    await expect(getCleanupPreview()).resolves.toEqual({
      deletable: 1,
      kept: 1,
    });

    const { keptComment, orphanComment } = installMixedColocationScenario();
    await expect(cleanupResolvedComments()).resolves.toEqual({
      deleted: 1,
      kept: 1,
    });

    expect(keptComment.delete).not.toHaveBeenCalled();
    expect(orphanComment.delete).toHaveBeenCalledOnce();
  });

  it("treats Stylistic-shaped comments by content, not by author name", async () => {
    const stylisticComment = makeComment(
      "Usuario de prueba",
      undefined,
      "[gramática]\nRedundancia pronominal.",
    );

    await installWordCleanupContext({
      comments: [stylisticComment],
      contentControls: [],
    });

    await expect(cleanupResolvedComments()).resolves.toEqual({
      deleted: 1,
      kept: 0,
    });
    expect(stylisticComment.delete).toHaveBeenCalledOnce();
  });

  it("propagates Word.run errors instead of swallowing them", async () => {
    const failure = new Error("Word sync failed");
    const { context } = installWordCleanupContext({
      comments: [makeComment("Usuario de prueba", undefined, "[Claridad]\nMas claro")],
      contentControls: [],
    });
    context.sync.mockRejectedValueOnce(failure);

    await expect(cleanupResolvedComments()).rejects.toThrow("Word sync failed");
  });

  it("exports the overlap allowlist used by cleanup colocation checks", () => {
    expect(OVERLAPPING_RELATIONS).toEqual(
      expect.arrayContaining([
        "Equal",
        "Contains",
        "Inside",
        "OverlapsBefore",
        "OverlapsAfter",
      ]),
    );
  });
});
