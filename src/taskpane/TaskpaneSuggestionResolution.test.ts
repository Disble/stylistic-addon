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
      liItems[0].querySelector('[data-action="accept"]')?.getAttribute(
        "data-suggestion-id",
      ),
    ).toBe("s-1");
    expect(
      liItems[0].querySelector('[data-action="reject"]')?.getAttribute(
        "data-suggestion-id",
      ),
    ).toBe("s-1");
    expect(
      liItems[1].querySelector('[data-action="accept"]')?.getAttribute(
        "data-suggestion-id",
      ),
    ).toBe("s-2");
  });

  it("removes action buttons for failed suggestions", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-fail" });

    const liItems = await renderViaEmitter(doc, [suggestion], ["s-fail"]);

    expect(liItems[0].querySelector('[data-action="accept"]')).toBeNull();
    expect(liItems[0].querySelector('[data-action="reject"]')).toBeNull();
    expect((liItems[0].querySelector(".result-failed") as FakeElement).textContent).toBe(
      `No encontrado: "${suggestion.anchor}"`,
    );
  });

  it("clicking Accept applies terminal accepted UI and removes buttons", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });
    taskpaneMocks.getCleanupPreview
      .mockResolvedValueOnce({ deletable: 0, kept: 1 })
      .mockResolvedValueOnce({ deletable: 0, kept: 1 })
      .mockResolvedValueOnce({ deletable: 1, kept: 0 });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.acceptSuggestion).toHaveBeenCalledWith(suggestion);
    expect(li.classList.contains("result-accepted")).toBe(true);
    expect(li.querySelector('[data-action="accept"]')).toBeNull();
    expect(li.querySelector('[data-action="reject"]')).toBeNull();
    expect(doc.getElementById("cleanup-section")!.style.display).toBe("block");
  });

  it("clicking Reject applies terminal rejected UI and removes buttons", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });
    taskpaneMocks.getCleanupPreview
      .mockResolvedValueOnce({ deletable: 0, kept: 1 })
      .mockResolvedValueOnce({ deletable: 0, kept: 1 })
      .mockResolvedValueOnce({ deletable: 1, kept: 0 });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const rejectBtn = li.querySelector('[data-action="reject"]') as FakeElement;

    rejectBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.rejectSuggestion).toHaveBeenCalledWith(suggestion);
    expect(li.classList.contains("result-rejected")).toBe(true);
    expect(li.querySelector('[data-action="accept"]')).toBeNull();
    expect(li.querySelector('[data-action="reject"]')).toBeNull();
    expect(doc.getElementById("cleanup-section")!.style.display).toBe("block");
  });

  it("keeps terminal rejected UI even when adapter ignored late cleanup failure", async () => {
    taskpaneMocks.rejectSuggestion.mockResolvedValue({
      status: "rejected",
      trackedChangesAffected: 2,
      commentDeleted: false,
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({
      id: "s-1",
      anchor: "con la Jing",
      suggestedText: "con Jing",
    });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const rejectBtn = li.querySelector('[data-action="reject"]') as FakeElement;

    rejectBtn.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(li.classList.contains("result-rejected")).toBe(true);
    expect(li.querySelector(".result-actions")).toBeNull();
    expect(li.textContent).not.toContain("(aplicación falló)");
  });

  it("marks already-resolved accept as terminal and sends positive feedback", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "already-resolved",
      trackedChangesAffected: 0,
      commentDeleted: false,
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

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
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const rejectBtn = li.querySelector('[data-action="reject"]') as FakeElement;

    rejectBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

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
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();

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
    }>();
    taskpaneMocks.acceptSuggestion.mockReturnValue(firstCall);

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    acceptBtn.click();
    await Promise.resolve();

    resolve({
      status: "accepted",
      trackedChangesAffected: 1,
      commentDeleted: false,
    });
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.acceptSuggestion).toHaveBeenCalledTimes(1);
    expect(li.classList.contains("result-accepted")).toBe(true);
  });

  it("allows retry after an error and reaches accepted state on the second click", async () => {
    taskpaneMocks.acceptSuggestion
      .mockResolvedValueOnce({
        status: "error",
        trackedChangesAffected: 0,
        commentDeleted: false,
        error: "timeout",
      })
      .mockResolvedValueOnce({
        status: "accepted",
        trackedChangesAffected: 1,
        commentDeleted: true,
      });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();
    expect(acceptBtn.disabled).toBe(false);

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(li.classList.contains("result-accepted")).toBe(true);
    expect(taskpaneMocks.acceptSuggestion).toHaveBeenCalledTimes(2);
  });

  it("re-enables buttons and surfaces status text on error", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "error",
      trackedChangesAffected: 0,
      commentDeleted: false,
      error: "El documento está protegido",
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(acceptBtn.disabled).toBe(false);
    expect(
      (li.querySelector('[data-action="reject"]') as FakeElement).disabled,
    ).toBe(false);
    expect(doc.getElementById("status-bar")?.textContent).toBe(
      "El documento está protegido",
    );
  });

  it("re-enables buttons when the adapter reports not-found", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValue({
      status: "not-found",
      trackedChangesAffected: 0,
      commentDeleted: false,
      error: "Texto no encontrado",
    });

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const acceptBtn = li.querySelector('[data-action="accept"]') as FakeElement;

    acceptBtn.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(acceptBtn.disabled).toBe(false);
    expect(
      (li.querySelector('[data-action="reject"]') as FakeElement).disabled,
    ).toBe(false);
  });
});
