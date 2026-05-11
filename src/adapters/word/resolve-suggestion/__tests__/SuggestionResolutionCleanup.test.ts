import { beforeEach, describe, expect, it, vi } from "vitest";
import { SuggestionResolutionCleanup } from "../SuggestionResolutionCleanup";
import type { ColocatedCommentContext } from "../ResolutionContext";

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
  const context = {
    document: {
      changeTrackingMode: "trackAll",
      contentControls: { items: [], load: vi.fn() },
      load: vi.fn(),
    },
    sync: syncMock,
  } as unknown as Word.RequestContext;
  return { context, syncMock };
}

type CleanupGraphTrackedChange = {
  id: string;
  source: "stylistic" | "foreign-review";
  accept: ReturnType<typeof vi.fn>;
  reject: ReturnType<typeof vi.fn>;
};

type CleanupGraphContentControl = {
  id: string;
  role: "operational-wrapper" | "inserted-side" | "foreign";
  tag: string;
  deleted: boolean;
  modeAtDelete?: string;
  delete: ReturnType<typeof vi.fn>;
};

type CleanupGraphComment = {
  id: string;
  label: string;
  authorName: string;
  content: string;
  deleted: boolean;
  delete: ReturnType<typeof vi.fn>;
};

type CleanupDocumentGraph = {
  comments: CleanupGraphComment[];
  contentControls: CleanupGraphContentControl[];
  trackedChanges: CleanupGraphTrackedChange[];
};

type IntegratedCleanupHarness = {
  context: Word.RequestContext;
  graph: CleanupDocumentGraph;
  colocatedComment: ColocatedCommentContext;
  locatedCommentNode: CleanupGraphComment;
  preservedCommentNode: CleanupGraphComment;
  foreignCommentNode: CleanupGraphComment;
  wrapperContentControl: CleanupGraphContentControl;
  insertedSideContentControl: CleanupGraphContentControl;
  foreignContentControl: CleanupGraphContentControl;
  foreignTrackedChange: CleanupGraphTrackedChange;
};

/**
 * Builds a shared document graph so cleanup assertions are anchored to the same
 * fixture state instead of detached spies. The cleanup method only receives the
 * selected comment proxy, but the surrounding graph lets the test verify that
 * comments are the only mutated artifact kind inside this repo's Office mock.
 */
function buildIntegratedCleanupHarness(): IntegratedCleanupHarness {
  const { context } = buildContextStub();
  const graph: CleanupDocumentGraph = {
    comments: [],
    contentControls: [],
    trackedChanges: [],
  };

  /** Removes one comment from the shared graph. */
  const deleteCommentFromGraph = (commentId: string) => {
    const comment = graph.comments.find((candidate) => candidate.id === commentId);
    if (!comment) {
      return;
    }

    comment.deleted = true;
    graph.comments = graph.comments.filter((candidate) => candidate.id !== commentId);
  };

  const wrapperContentControl: CleanupGraphContentControl = {
    id: "cc-wrapper-1",
    role: "operational-wrapper",
    tag: "stylistic-operational-wrapper:s-cleanup",
    deleted: false,
    delete: vi.fn(() => {
      wrapperContentControl.modeAtDelete = context.document.changeTrackingMode;
      wrapperContentControl.deleted = true;
      graph.contentControls = graph.contentControls.filter(
        (candidate) => candidate.id !== wrapperContentControl.id
      );
    }),
  };

  const insertedSideContentControl: CleanupGraphContentControl = {
    id: "cc-inserted-1",
    role: "inserted-side",
    tag: "stylistic:track-change:s-cleanup",
    deleted: false,
    delete: vi.fn(() => {
      insertedSideContentControl.modeAtDelete = context.document.changeTrackingMode;
      insertedSideContentControl.deleted = true;
      graph.contentControls = graph.contentControls.filter(
        (candidate) => candidate.id !== insertedSideContentControl.id
      );
    }),
  };

  const foreignContentControl: CleanupGraphContentControl = {
    id: "cc-foreign-1",
    role: "foreign",
    tag: "stylistic:track-change:s-other",
    deleted: false,
    delete: vi.fn(() => {
      foreignContentControl.modeAtDelete = context.document.changeTrackingMode;
      foreignContentControl.deleted = true;
      graph.contentControls = graph.contentControls.filter(
        (candidate) => candidate.id !== foreignContentControl.id
      );
    }),
  };

  const foreignTrackedChange: CleanupGraphTrackedChange = {
    id: "tc-foreign-1",
    source: "foreign-review",
    accept: vi.fn(),
    reject: vi.fn(),
  };

  const locatedCommentNode: CleanupGraphComment = {
    id: "comment-target",
    label: "resolved-stylistic-comment",
    authorName: "Usuario de prueba",
    content: "[Claridad]\nTexto mas claro.",
    deleted: false,
    delete: vi.fn(() => deleteCommentFromGraph("comment-target")),
  };

  const preservedCommentNode: CleanupGraphComment = {
    id: "comment-preserved",
    label: "other-stylistic-comment",
    authorName: "Usuario de prueba",
    content: "[Registro]\nSegunda sugerencia pendiente.",
    deleted: false,
    delete: vi.fn(() => deleteCommentFromGraph("comment-preserved")),
  };

  const foreignCommentNode: CleanupGraphComment = {
    id: "comment-foreign",
    label: "foreign-comment",
    authorName: "Reviewer",
    content: "Comentario ajeno al addon.",
    deleted: false,
    delete: vi.fn(() => deleteCommentFromGraph("comment-foreign")),
  };

  graph.comments = [locatedCommentNode, preservedCommentNode, foreignCommentNode];
  graph.contentControls = [
    wrapperContentControl,
    insertedSideContentControl,
    foreignContentControl,
  ];
  graph.trackedChanges = [foreignTrackedChange];

  const contentControlsCollection = {
    load: vi.fn(),
    get items() {
      return graph.contentControls;
    },
  };
  Object.assign(context.document, {
    contentControls: contentControlsCollection,
  });

  return {
    context,
    graph,
    colocatedComment: {
      comment: {
        delete: locatedCommentNode.delete,
      } as unknown as Word.Comment,
      range: {} as unknown as Word.Range,
    },
    locatedCommentNode,
    preservedCommentNode,
    foreignCommentNode,
    wrapperContentControl,
    insertedSideContentControl,
    foreignContentControl,
    foreignTrackedChange,
  };
}

function makeColocatedComment(deleteImpl: () => void = () => undefined): ColocatedCommentContext {
  return {
    comment: { delete: vi.fn(deleteImpl) } as unknown as Word.Comment,
    range: {} as unknown as Word.Range,
  };
}

describe("SuggestionResolutionCleanup soft-success contract", () => {
  beforeEach(() => {
    (globalThis as unknown as { Word: { ChangeTrackingMode: Record<string, string> } }).Word = {
      ChangeTrackingMode: {
        off: "off",
        trackAll: "trackAll",
        trackMine: "trackMine",
      },
    };
  });

  it("returns commentDeleted=true when comment.delete()/sync raises GeneralException after the host already invalidated the comment proxy", async () => {
    const cleanup = new SuggestionResolutionCleanup("s-1", "reject");
    const { context, syncMock } = buildContextStub();
    syncMock.mockImplementationOnce(async () => {
      throw new Error("GeneralException");
    });
    const colocatedComment = makeColocatedComment();

    const commentDeleted = await cleanup.deleteLocatedStylisticComment(context, colocatedComment);

    expect(commentDeleted).toBe(true);
  });

  it("propagates non-GeneralException errors from comment.delete()/sync so legitimate failures still abort the workflow", async () => {
    const cleanup = new SuggestionResolutionCleanup("s-2", "reject");
    const { context, syncMock } = buildContextStub();
    syncMock.mockImplementationOnce(async () => {
      throw new Error("InvalidOperationInCellEdit");
    });
    const colocatedComment = makeColocatedComment();

    await expect(cleanup.deleteLocatedStylisticComment(context, colocatedComment)).rejects.toThrow(
      "InvalidOperationInCellEdit"
    );
  });

  it("exposes explicit metadata cleanup for resolved track-change wrappers", () => {
    const cleanup = new SuggestionResolutionCleanup("s-3", "reject");

    expect("deleteResolvedTrackChangeMetadata" in cleanup).toBe(true);
  });

  it("deletes only the located Stylistic comment from a shared document graph while preserving wrapper CCs, inserted-side CCs, and foreign tracked changes", async () => {
    const cleanup = new SuggestionResolutionCleanup("s-preserve", "accept");
    const harness = buildIntegratedCleanupHarness();

    const commentDeleted = await cleanup.deleteLocatedStylisticCommentAfterResolution(
      harness.context,
      harness.colocatedComment
    );

    expect(commentDeleted).toBe(true);
    expect(harness.colocatedComment.comment.delete).toHaveBeenCalledOnce();
    expect(harness.locatedCommentNode.deleted).toBe(true);
    expect(harness.graph.comments.map((comment) => comment.id)).toEqual([
      harness.preservedCommentNode.id,
      harness.foreignCommentNode.id,
    ]);
    expect(harness.preservedCommentNode.deleted).toBe(false);
    expect(harness.foreignCommentNode.deleted).toBe(false);
    expect(harness.graph.contentControls.map((cc) => cc.id)).toEqual([
      harness.wrapperContentControl.id,
      harness.insertedSideContentControl.id,
      harness.foreignContentControl.id,
    ]);
    expect(harness.wrapperContentControl.deleted).toBe(false);
    expect(harness.insertedSideContentControl.deleted).toBe(false);
    expect(harness.foreignContentControl.deleted).toBe(false);
    expect(harness.wrapperContentControl.delete).not.toHaveBeenCalled();
    expect(harness.insertedSideContentControl.delete).not.toHaveBeenCalled();
    expect(harness.foreignContentControl.delete).not.toHaveBeenCalled();
    expect(harness.graph.trackedChanges).toEqual([harness.foreignTrackedChange]);
    expect(harness.foreignTrackedChange.accept).not.toHaveBeenCalled();
    expect(harness.foreignTrackedChange.reject).not.toHaveBeenCalled();
  });

  it("deletes exact resolved track-change metadata with Track Changes temporarily disabled and visible text preserved", async () => {
    const cleanup = new SuggestionResolutionCleanup("s-cleanup", "accept");
    const harness = buildIntegratedCleanupHarness();

    const result = await cleanup.deleteResolvedTrackChangeMetadata(harness.context);

    expect(result.deletedContentControls).toEqual([
      harness.wrapperContentControl.tag,
      harness.insertedSideContentControl.tag,
    ]);
    expect(result.failedContentControls).toEqual([]);
    expect(harness.wrapperContentControl.delete).toHaveBeenCalledWith(true);
    expect(harness.insertedSideContentControl.delete).toHaveBeenCalledWith(true);
    expect(harness.wrapperContentControl.modeAtDelete).toBe("off");
    expect(harness.insertedSideContentControl.modeAtDelete).toBe("off");
    expect(harness.context.document.changeTrackingMode).toBe("trackAll");
    expect(harness.graph.contentControls.map((cc) => cc.id)).toEqual([
      harness.foreignContentControl.id,
    ]);
    expect(harness.foreignContentControl.delete).not.toHaveBeenCalled();
    expect(harness.graph.comments).toHaveLength(3);
    expect(harness.graph.trackedChanges).toEqual([harness.foreignTrackedChange]);
  });

  it("treats absent inserted-side metadata as an idempotent resolved reject cleanup", async () => {
    const cleanup = new SuggestionResolutionCleanup("s-cleanup", "reject");
    const harness = buildIntegratedCleanupHarness();
    harness.graph.contentControls = harness.graph.contentControls.filter(
      (cc) => cc.id !== harness.insertedSideContentControl.id
    );

    const result = await cleanup.deleteResolvedTrackChangeMetadata(harness.context);

    expect(result.deletedContentControls).toEqual([harness.wrapperContentControl.tag]);
    expect(result.failedContentControls).toEqual([]);
    expect(harness.wrapperContentControl.delete).toHaveBeenCalledWith(true);
    expect(harness.insertedSideContentControl.delete).not.toHaveBeenCalled();
    expect(harness.context.document.changeTrackingMode).toBe("trackAll");
    expect(harness.graph.contentControls.map((cc) => cc.id)).toEqual([
      harness.foreignContentControl.id,
    ]);
  });
});
