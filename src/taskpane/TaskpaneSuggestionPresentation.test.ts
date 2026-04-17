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

describe("taskpane suggestion presentation", () => {
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

  it("renders comment-only suggestions without diff blocks and with text action labels", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({
      id: "s-co",
      type: "comment-only",
      suggestedText: undefined,
    });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];

    expect(li.querySelector(".card-diff")).toBeNull();
    expect(li.querySelector('[data-action="accept"]')?.textContent).toBe(
      "Entendido",
    );
    expect(li.querySelector('[data-action="reject"]')?.textContent).toBe(
      "Ignorar",
    );
    expect(li.querySelector(".result-type-badge--comment")?.textContent).toBe(
      "comentario",
    );
  });

  it("renders track-change suggestions with diff blocks and symbolic action labels", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({
      id: "s-tc",
      type: "track-change",
      suggestedText: "texto sugerido",
    });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];

    expect(li.querySelector(".card-diff")).not.toBeNull();
    expect(li.querySelector('[data-action="accept"]')?.textContent).toBe("✓");
    expect(li.querySelector('[data-action="reject"]')?.textContent).toBe("✗");
  });

  it("uses the suggestion identity for clickable navigation", async () => {
    const doc = createTaskpaneDocument();
    const suggestion = makeSuggestion({
      id: "s-nav",
      anchor: "fragmento exacto",
      context: "Un contexto con fragmento exacto adentro.",
    });

    const li = (await renderViaEmitter(doc, [suggestion]))[0];
    const clickable = li.querySelector(".card-clickable-area") as FakeElement;

    expect((li.querySelector(".result-original") as FakeElement).textContent).toBe(
      "fragmento exacto",
    );
    clickable.click();

    expect(taskpaneMocks.navigateToText).toHaveBeenCalledWith(suggestion);
  });
});
