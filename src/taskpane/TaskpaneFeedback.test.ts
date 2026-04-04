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

/** Flushes taskpane microtasks for async review resolution tests. */
async function flushTaskpaneWork(times = 8) {
  for (let index = 0; index < times; index += 1) {
    await Promise.resolve();
  }
}

describe("taskpane feedback controls", () => {
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

  it("renders a feedback button and accordion for each non-failed suggestion", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];

    expect(li.querySelector('[data-action="feedback"]')?.getAttribute("aria-label")).toBe(
      "Dejar feedback",
    );
    expect(li.querySelector(".feedback-accordion")).not.toBeNull();
    expect(li.querySelector(".feedback-textarea")).not.toBeNull();
  });

  it("omits feedback controls for failed suggestions", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-fail" });

    const li = (await renderViaEmitter(doc, [suggestion], ["s-fail"]))[0];

    expect(li.querySelector('[data-action="feedback"]')).toBeNull();
    expect(li.querySelector(".feedback-accordion")).toBeNull();
  });

  it("toggles the feedback accordion when the feedback button is clicked", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const feedbackBtn = li.querySelector('[data-action="feedback"]') as FakeElement;
    const accordion = li.querySelector(".feedback-accordion") as FakeElement;

    expect(accordion.classList.contains("feedback-accordion--open")).toBe(false);
    feedbackBtn.click();
    expect(accordion.classList.contains("feedback-accordion--open")).toBe(true);
    feedbackBtn.click();
    expect(accordion.classList.contains("feedback-accordion--open")).toBe(false);
  });

  it("sends positive feedback payloads after accept", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({
      id: "s-1",
      category: "Redundancia",
      anchor: "completamente necesario",
      context: "Frase con completamente necesario.",
      suggestedText: "necesario",
      justification: "Ya implica completitud.",
      severity: "high",
    });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    (li.querySelector('[data-action="accept"]') as FakeElement).click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.feedbackSendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({
        rating: "positive",
        category: "Redundancia",
        originalText: "completamente necesario",
        suggestedText: "necesario",
        justification: "Ya implica completitud.",
        severity: "high",
      }),
    );
  });

  it("sends negative feedback payloads after reject", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({
      id: "s-1",
      category: "Muletilla",
      anchor: "básicamente",
      context: "Frase con básicamente.",
      suggestedText: "",
      justification: "Frase de relleno.",
      severity: "medium",
    });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    (li.querySelector('[data-action="reject"]') as FakeElement).click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.feedbackSendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({
        rating: "negative",
        category: "Muletilla",
        originalText: "básicamente",
        suggestedText: "",
        justification: "Frase de relleno.",
        severity: "medium",
      }),
    );
  });

  it("keeps feedback non-blocking even when the adapter rejects it later", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValueOnce({
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
    });
    taskpaneMocks.feedbackSendFeedback.mockRejectedValueOnce(
      new Error("feedback failed"),
    );

    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-feedback" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    (li.querySelector('[data-action="accept"]') as FakeElement).click();
    await flushTaskpaneWork();

    expect(li.classList.contains("result-accepted")).toBe(true);
  });

  it("does not send feedback when the resolution ends in identity-lost", async () => {
    taskpaneMocks.acceptSuggestion.mockResolvedValueOnce({
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
    (li.querySelector('[data-action="accept"]') as FakeElement).click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.feedbackSendFeedback).not.toHaveBeenCalled();
  });

  it("omits empty textarea comments from the feedback payload", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    (li.querySelector(".feedback-textarea") as FakeElement).value = "";
    (li.querySelector('[data-action="accept"]') as FakeElement).click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.feedbackSendFeedback.mock.calls[0][0]).not.toHaveProperty(
      "comment",
    );
  });

  it("includes non-empty textarea comments and justification in the feedback payload", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({ id: "s-1", justification: "Es más claro" });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    (li.querySelector(".feedback-textarea") as FakeElement).value = "Muy buen cambio";
    (li.querySelector('[data-action="accept"]') as FakeElement).click();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(taskpaneMocks.feedbackSendFeedback).toHaveBeenCalledWith(
      expect.objectContaining({
        comment: "Muy buen cambio",
        justification: "Es más claro",
      }),
    );
  });
});
