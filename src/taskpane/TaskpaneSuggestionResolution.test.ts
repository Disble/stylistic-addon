import type { SuggestionResolutionMediatorResult } from "../domain/types";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  createTaskpaneDocument,
  FakeElement,
  getTaskpaneMocks,
  makeSuggestion,
  renderViaEmitter,
  resetTaskpaneHarness,
  teardownTaskpaneHarness,
} from "./TaskpaneTestHelper";

/** Returns a required fake DOM element by id. */
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

/** Returns a required fake child selected from the given parent. */
function getRequiredChild(parent: FakeElement, selector: string): FakeElement {
  const child = parent.querySelector(selector);
  if (!child) {
    throw new Error(`Missing child for selector: ${selector}`);
  }

  return child;
}

/** Flushes queued microtasks from taskpane action handlers. */
async function flushTaskpaneWork(times = 8): Promise<void> {
  for (let index = 0; index < times; index += 1) {
    await Promise.resolve();
  }
}

/** Builds a compact mediator result fixture for taskpane action tests. */
function makeMediatorResult(
  overrides: Partial<SuggestionResolutionMediatorResult> = {},
): SuggestionResolutionMediatorResult {
  return {
    status: "accepted",
    trackedChangesAffected: 2,
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
    ...overrides,
  };
}

describe("TaskpaneSuggestionResolution", () => {
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

  it("marks an accepted suggestion as terminal and updates taskpane CTAs", async () => {
    taskpaneMocks.getCleanupPreview.mockResolvedValue({
      deletable: 2,
      kept: 0,
    });
    taskpaneMocks.acceptSuggestion.mockResolvedValueOnce(
      makeMediatorResult({
        status: "accepted",
        pendingAfter: {
          pendingStylisticArtifacts: 0,
          hasPendingStylisticArtifacts: false,
          trackChangesActive: true,
        },
        documentState: "ready-to-disable-track-changes",
      }),
    );

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-accept" });
    const li = (await renderViaEmitter(doc, [suggestion]))[0];

    getRequiredChild(li, '[data-action="accept"]').click();
    await flushTaskpaneWork();

    expect(taskpaneMocks.acceptSuggestion).toHaveBeenCalledWith(suggestion);
    expect(li.classList.contains("result-accepted")).toBe(true);
    expect(li.querySelector(".result-actions")).toBeNull();
    expect(getRequiredElement(doc, "cleanup-section").style.display).toBe("block");
    expect(getRequiredElement(doc, "disable-track-changes-section").style.display).toBe("block");
  });

  it("marks a rejected suggestion as terminal and removes action buttons", async () => {
    taskpaneMocks.rejectSuggestion.mockResolvedValueOnce(
      makeMediatorResult({
        status: "rejected",
        feedbackStatus: "sent",
      }),
    );

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-reject" });
    const li = (await renderViaEmitter(doc, [suggestion]))[0];

    getRequiredChild(li, '[data-action="reject"]').click();
    await flushTaskpaneWork();

    expect(taskpaneMocks.rejectSuggestion).toHaveBeenCalledWith(suggestion);
    expect(li.classList.contains("result-rejected")).toBe(true);
    expect(li.querySelector(".result-actions")).toBeNull();
  });

  it("re-enables buttons and keeps the card non-terminal on unobservable", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValueOnce(
      makeMediatorResult({
        status: "unobservable",
        trackedChangesAffected: 0,
        commentDeleted: false,
        feedbackStatus: "skipped",
        error:
          "Word no expuso suficiente evidencia operacional para confirmar la resolución.",
      }),
    );

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-unobservable" });
    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = getRequiredChild(li, '[data-action="accept"]');
    const rejectBtn = getRequiredChild(li, '[data-action="reject"]');

    acceptBtn.click();
    await flushTaskpaneWork();

    expect(acceptBtn.disabled).toBe(false);
    expect(rejectBtn.disabled).toBe(false);
    expect(li.classList.contains("result-accepted")).toBe(false);
    expect(li.classList.contains("result-unobservable")).toBe(false);
    expect(getRequiredElement(doc, "status-bar").textContent).toBe(
      "Word no expuso suficiente evidencia operacional para confirmar la resolución.",
    );
    expect(taskpaneMocks.feedbackSendFeedback).not.toHaveBeenCalled();
  });

  it("renders ambiguous-location as terminal manual-review UI", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValueOnce(
      makeMediatorResult({
        status: "ambiguous-location",
        trackedChangesAffected: 0,
        commentDeleted: false,
        feedbackStatus: "skipped",
        error: "La ubicación de la sugerencia es ambigua.",
      }),
    );

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-ambiguous" });
    const li = (await renderViaEmitter(doc, [suggestion]))[0];

    getRequiredChild(li, '[data-action="accept"]').click();
    await flushTaskpaneWork();

    expect(li.classList.contains("result-ambiguous-location")).toBe(true);
    expect(li.querySelector(".result-actions")).toBeNull();
    expect(li.querySelector(".result-ambiguous-location-note")?.textContent).toBe(
      "(resolución ambigua; reanalizá la sugerencia)",
    );
    expect(taskpaneMocks.feedbackSendFeedback).not.toHaveBeenCalled();
  });

  it("renders mixed-group as terminal manual-review UI", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValueOnce(
      makeMediatorResult({
        status: "mixed-group",
        trackedChangesAffected: 0,
        commentDeleted: false,
        feedbackStatus: "skipped",
        error: "El grupo contiguo requiere resolución grupal coherente.",
      }),
    );

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-mixed" });
    const li = (await renderViaEmitter(doc, [suggestion]))[0];

    getRequiredChild(li, '[data-action="accept"]').click();
    await flushTaskpaneWork();

    expect(li.classList.contains("result-mixed-group")).toBe(true);
    expect(li.querySelector(".result-actions")).toBeNull();
    expect(li.querySelector(".result-mixed-group-note")?.textContent).toBe(
      "(resolución ambigua; reanalizá la sugerencia)",
    );
    expect(taskpaneMocks.feedbackSendFeedback).not.toHaveBeenCalled();
  });
});
