import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  createTaskpaneDocument,
  deferred,
  FakeElement,
  getTaskpaneMocks,
  makeSuggestion,
  renderViaEmitter,
  resetTaskpaneHarness,
  teardownTaskpaneHarness,
} from "./TaskpaneTestHelper";

/** Returns a required fake document element or throws in tests. */
function getRequiredElement(
  doc: ReturnType<typeof createTaskpaneDocument>,
  id: string,
): FakeElement {
  const element = doc.getElementById(id);
  if (!element) {
    throw new Error(`Missing fake DOM element: ${id}`);
  }
  return element;
}

/** Returns a required fake child element from a selector or throws in tests. */
function getRequiredChild(parent: FakeElement, selector: string): FakeElement {
  const child = parent.querySelector(selector);
  if (!child) {
    throw new Error(`Missing child for selector: ${selector}`);
  }
  return child;
}

/** Flushes microtasks scheduled by the taskpane review mediator/workflows. */
async function flushTaskpaneWork(times = 8) {
  for (let index = 0; index < times; index += 1) {
    await Promise.resolve();
  }
}

describe("taskpane suggestion resolution guardrails", () => {
  const taskpaneMocks = getTaskpaneMocks();
  let logSpy: ReturnType<typeof vi.spyOn>;
  let warnSpy: ReturnType<typeof vi.spyOn>;
  let errorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    resetTaskpaneHarness();
    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
    warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    errorSpy = vi.spyOn(console, "error").mockImplementation(() => {});
  });

  afterEach(() => {
    logSpy.mockRestore();
    warnSpy.mockRestore();
    errorSpy.mockRestore();
    teardownTaskpaneHarness();
  });

  it("renders accept and reject buttons for non-failed suggestions", async () => {
    const doc = createTaskpaneDocument();
    const s1 = makeSuggestion({ id: "s-1" });
    const s2 = makeSuggestion({ id: "s-2" });

    const liItems = await renderViaEmitter(doc, [s1, s2]);

    expect(liItems).toHaveLength(2);
    expect(
      liItems[0].querySelector('[data-action="accept"]')?.dataset.suggestionId,
    ).toBe("s-1");
    expect(
      liItems[0].querySelector('[data-action="reject"]')?.dataset.suggestionId,
    ).toBe("s-1");
    expect(
      liItems[1].querySelector('[data-action="accept"]')?.dataset.suggestionId,
    ).toBe("s-2");
  });

  it("removes action buttons for failed suggestions", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-fail" });

    const liItems = await renderViaEmitter(doc, [suggestion], ["s-fail"]);

    expect(liItems[0].querySelector('[data-action="accept"]')).toBeNull();
    expect(liItems[0].querySelector('[data-action="reject"]')).toBeNull();
    expect(getRequiredChild(liItems[0], ".result-failed").textContent).toBe(
      `No encontrado: "${suggestion.anchor}"`,
    );
  });

  it("renders non-not-found failures without lying about the cause", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-fail" });

    taskpaneMocks.run.mockImplementationOnce(async (ctx: any) => {
      ctx.emitter.emitComplete(
        [suggestion],
        {
          successCount: 0,
          failedSuggestions: [
            {
              suggestion,
              reason: "command-error",
              message: "GeneralException",
            },
          ],
          pendingAfter: {
            pendingStylisticArtifacts: 0,
            hasPendingStylisticArtifacts: false,
            trackChangesActive: false,
          },
          documentState: "idle",
          trackChangesActivatedForBatch: false,
        },
        [],
        false,
      );
    });

    const liItems = await renderViaEmitter(doc, [suggestion]);

    expect(getRequiredChild(liItems[0], ".result-failed").textContent).toBe(
      `No se pudo aplicar: "${suggestion.anchor}"`,
    );
    expect(
      getRequiredChild(liItems[0], ".result-failure-detail").textContent,
    ).toBe("GeneralException");
  });

  it("clicking Accept applies terminal accepted UI and removes buttons", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });
    taskpaneMocks.getCleanupPreview
      .mockResolvedValueOnce({ deletable: 0, kept: 1 })
      .mockResolvedValueOnce({ deletable: 0, kept: 1 })
      .mockResolvedValueOnce({ deletable: 2, kept: 0 });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');

    acceptBtn.click();
    await flushTaskpaneWork();

    expect(taskpaneMocks.acceptSuggestion).toHaveBeenCalledWith(suggestion);
    expect(li.classList.contains("result-accepted")).toBe(true);
    expect(li.querySelector('[data-action="accept"]')).toBeNull();
    expect(li.querySelector('[data-action="reject"]')).toBeNull();
    expect(getRequiredElement(doc, "cleanup-section").style.display).toBe("block");
    expect(
      getRequiredElement(doc, "disable-track-changes-section").style.display,
    ).toBe("none");
  });

  it("clicking Reject applies terminal rejected UI and removes buttons", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });
    taskpaneMocks.getCleanupPreview
      .mockResolvedValueOnce({ deletable: 0, kept: 1 })
      .mockResolvedValueOnce({ deletable: 0, kept: 1 })
      .mockResolvedValueOnce({ deletable: 2, kept: 0 });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const rejectBtn = getRequiredChild(li, '[data-action="reject"]');

    rejectBtn.click();
    await flushTaskpaneWork();

    expect(taskpaneMocks.rejectSuggestion).toHaveBeenCalledWith(suggestion);
    expect(li.classList.contains("result-rejected")).toBe(true);
    expect(li.querySelector('[data-action="accept"]')).toBeNull();
    expect(li.querySelector('[data-action="reject"]')).toBeNull();
    expect(getRequiredElement(doc, "cleanup-section").style.display).toBe("block");
    expect(
      getRequiredElement(doc, "disable-track-changes-section").style.display,
    ).toBe("none");
  });

  it("keeps terminal rejected UI even when adapter ignored late cleanup failure", async () => {
    taskpaneMocks.rejectSuggestion.mockResolvedValue({
      status: "rejected",
      trackedChangesAffected: 2,
      commentDeleted: false,
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({
      id: "s-1",
      anchor: "con la Jing",
      suggestedText: "con Jing",
    });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const rejectBtn = getRequiredChild(li, '[data-action="reject"]');

    rejectBtn.click();
    await flushTaskpaneWork();

    expect(li.classList.contains("result-rejected")).toBe(true);
    expect(li.querySelector(".result-actions")).toBeNull();
    expect(li.textContent).not.toContain("(aplicación falló)");
  });

  it("marks already-resolved accept as terminal and sends positive feedback", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "already-resolved",
      trackedChangesAffected: 0,
      commentDeleted: false,
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      feedbackStatus: "sent",
      taskpaneState: {
        documentState: "pending-review",
        showDisableTrackChangesCta: false,
        showCleanupSection: false,
      },
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');

    acceptBtn.click();
    await flushTaskpaneWork();

    expect(li.classList.contains("result-already-resolved")).toBe(true);
    expect(
      li.querySelector(".result-already-resolved-note")?.textContent,
    ).toBe("(ya resuelto)");
    expect(taskpaneMocks.feedbackSendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({ rating: "positive", originalText: suggestion.anchor }),
    );
  });

  it("marks already-resolved reject as terminal and sends negative feedback", async () => {
    taskpaneMocks.rejectSuggestion.mockResolvedValue({
      status: "already-resolved",
      trackedChangesAffected: 0,
      commentDeleted: false,
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      feedbackStatus: "sent",
      taskpaneState: {
        documentState: "pending-review",
        showDisableTrackChangesCta: false,
        showCleanupSection: false,
      },
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const rejectBtn = getRequiredChild(li, '[data-action="reject"]');

    rejectBtn.click();
    await flushTaskpaneWork();

    expect(li.classList.contains("result-already-resolved")).toBe(true);
    expect(taskpaneMocks.feedbackSendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({ rating: "negative", originalText: suggestion.anchor }),
    );
  });

  it("treats cc-not-found as terminal amber UI without sending feedback", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "cc-not-found",
      trackedChangesAffected: 0,
      commentDeleted: false,
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      feedbackStatus: "skipped",
      taskpaneState: {
        documentState: "pending-review",
        showDisableTrackChangesCta: false,
        showCleanupSection: false,
      },
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');

    acceptBtn.click();
    await flushTaskpaneWork();

    expect(li.classList.contains("result-cc-not-found")).toBe(true);
    expect(li.querySelector(".result-actions")).toBeNull();
    expect(li.querySelector(".result-cc-not-found-note")?.textContent).toBe(
      "(aplicación falló)",
    );
    expect(taskpaneMocks.feedbackSendFeedback).not.toHaveBeenCalled();
  });

  it("guards against double-clicking accept while resolution is still in flight", async () => {
    const { promise: firstCall, resolve } = deferred<{
      status: string;
      trackedChangesAffected: number;
      commentDeleted: boolean;
      pendingAfter: {
        pendingStylisticArtifacts: number;
        hasPendingStylisticArtifacts: boolean;
        trackChangesActive: boolean;
      };
      documentState: "pending-review" | "idle" | "ready-to-disable-track-changes";
      feedbackStatus?: string;
      taskpaneState?: {
        documentState: "pending-review" | "idle" | "ready-to-disable-track-changes";
        showDisableTrackChangesCta: boolean;
        showCleanupSection: boolean;
      };
    }>();
    taskpaneMocks.acceptSuggestion.mockReturnValue(firstCall);

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');

    acceptBtn.click();
    await flushTaskpaneWork(2);
    acceptBtn.click();
    await flushTaskpaneWork(2);

    resolve({
      status: "accepted",
      trackedChangesAffected: 1,
      commentDeleted: false,
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      feedbackStatus: "sent",
      taskpaneState: {
        documentState: "pending-review",
        showDisableTrackChangesCta: false,
        showCleanupSection: false,
      },
    });
    await flushTaskpaneWork();

    expect(taskpaneMocks.acceptSuggestion).toHaveBeenCalledTimes(1);
    expect(li.classList.contains("result-accepted")).toBe(true);
  });

  it("allows retry after an error and reaches accepted state on the second click", async () => {
    taskpaneMocks.acceptSuggestion
      .mockResolvedValueOnce({
        status: "error",
        trackedChangesAffected: 0,
        commentDeleted: false,
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        error: "timeout",
        feedbackStatus: "skipped",
        taskpaneState: {
          documentState: "pending-review",
          showDisableTrackChangesCta: false,
          showCleanupSection: false,
        },
      })
      .mockResolvedValueOnce({
        status: "accepted",
        trackedChangesAffected: 1,
        commentDeleted: true,
        pendingAfter: {
          pendingStylisticArtifacts: 1,
          hasPendingStylisticArtifacts: true,
          trackChangesActive: true,
        },
        documentState: "pending-review",
        feedbackStatus: "sent",
        taskpaneState: {
          documentState: "pending-review",
          showDisableTrackChangesCta: false,
          showCleanupSection: false,
        },
      });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');

    acceptBtn.click();
    await flushTaskpaneWork();
    expect(acceptBtn.disabled).toBe(false);

    acceptBtn.click();
    await flushTaskpaneWork();

    expect(li.classList.contains("result-accepted")).toBe(true);
    expect(taskpaneMocks.acceptSuggestion).toHaveBeenCalledTimes(2);
  });

  it("re-enables buttons and surfaces status text on error", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "error",
      trackedChangesAffected: 0,
      commentDeleted: false,
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      error: "El documento está protegido",
      feedbackStatus: "skipped",
      taskpaneState: {
        documentState: "pending-review",
        showDisableTrackChangesCta: false,
        showCleanupSection: false,
      },
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');

    acceptBtn.click();
    await flushTaskpaneWork();

    expect(acceptBtn.disabled).toBe(false);
    expect(
      getRequiredChild(li, '[data-action="reject"]').disabled,
    ).toBe(false);
    expect(doc.getElementById("status-bar")?.textContent).toBe(
      "El documento está protegido",
    );
  });

  it("re-enables buttons and does not mark terminal state on unobservable", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "unobservable",
      trackedChangesAffected: 0,
      commentDeleted: false,
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      error:
        "Word no expuso suficientes tracked changes para confirmar la resolución.",
      feedbackStatus: "skipped",
      taskpaneState: {
        documentState: "pending-review",
        showDisableTrackChangesCta: false,
        showCleanupSection: false,
      },
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');

    acceptBtn.click();
    await flushTaskpaneWork();

    expect(acceptBtn.disabled).toBe(false);
    expect(getRequiredChild(li, '[data-action="reject"]').disabled).toBe(false);
    expect(li.classList.contains("result-already-resolved")).toBe(false);
    expect(doc.getElementById("status-bar")?.textContent).toBe(
      "Word no expuso suficientes tracked changes para confirmar la resolución.",
    );
    expect(taskpaneMocks.feedbackSendFeedback).not.toHaveBeenCalled();
  });

  it("renders identity-lost as a terminal warning and skips feedback", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "identity-lost",
      trackedChangesAffected: 0,
      commentDeleted: false,
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      error: "La metadata compound-v2 de la sugerencia está incompleta o corrupta.",
      feedbackStatus: "skipped",
      taskpaneState: {
        documentState: "pending-review",
        showDisableTrackChangesCta: false,
        showCleanupSection: false,
      },
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-identity-lost" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');

    acceptBtn.click();
    await flushTaskpaneWork();

    expect(li.classList.contains("result-identity-lost")).toBe(true);
    expect(li.querySelector(".result-actions")).toBeNull();
    expect(li.querySelector(".result-identity-lost-note")?.textContent).toBe(
      "(metadata inconsistente; reanalizá la sugerencia)",
    );
    expect(taskpaneMocks.feedbackSendFeedback).not.toHaveBeenCalled();
  });

  it("re-enables buttons when the adapter reports not-found", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "not-found",
      trackedChangesAffected: 0,
      commentDeleted: false,
      pendingAfter: {
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      },
      documentState: "pending-review",
      error: "Texto no encontrado",
      feedbackStatus: "skipped",
      taskpaneState: {
        documentState: "pending-review",
        showDisableTrackChangesCta: false,
        showCleanupSection: false,
      },
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');

    acceptBtn.click();
    await flushTaskpaneWork();

    expect(acceptBtn.disabled).toBe(false);
    expect(
      getRequiredChild(li, '[data-action="reject"]').disabled,
    ).toBe(false);
  });

  it("shows the disable Track Changes CTA only when final resolution reaches zero pending", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "accepted",
      trackedChangesAffected: 1,
      commentDeleted: true,
      pendingAfter: {
        pendingStylisticArtifacts: 0,
        hasPendingStylisticArtifacts: false,
        trackChangesActive: true,
      },
      documentState: "ready-to-disable-track-changes",
      feedbackStatus: "sent",
      taskpaneState: {
        documentState: "ready-to-disable-track-changes",
        showDisableTrackChangesCta: true,
        showCleanupSection: false,
      },
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-final" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');

    acceptBtn.click();
    await flushTaskpaneWork();

    expect(
      getRequiredElement(doc, "disable-track-changes-section").style.display,
    ).toBe("block");
  });

  it("lets the user disable Track Changes from the explicit CTA", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-final" });
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "accepted",
      trackedChangesAffected: 1,
      commentDeleted: true,
      pendingAfter: {
        pendingStylisticArtifacts: 0,
        hasPendingStylisticArtifacts: false,
        trackChangesActive: true,
      },
      documentState: "ready-to-disable-track-changes",
      feedbackStatus: "sent",
      taskpaneState: {
        documentState: "ready-to-disable-track-changes",
        showDisableTrackChangesCta: true,
        showCleanupSection: false,
      },
    });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    getRequiredChild(li, '[data-action="accept"]').click();
    await flushTaskpaneWork();

    getRequiredElement(doc, "btn-disable-track-changes").click();
    await flushTaskpaneWork();

    expect(taskpaneMocks.disableTrackChanges).toHaveBeenCalledOnce();
    expect(
      getRequiredElement(doc, "disable-track-changes-section").style.display,
    ).toBe("none");
    expect(getRequiredElement(doc, "status-bar").textContent).toBe(
      "Control de cambios desactivado.",
    );
  });
});
