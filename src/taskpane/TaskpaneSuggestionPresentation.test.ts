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
import { buildResultsSummary } from "./SuggestionCardRenderer";

/** Returns a fake DOM element that the test requires to exist. */
function requireElement(container: FakeElement, selector: string): FakeElement {
  const element = container.querySelector(selector);
  expect(element).not.toBeNull();
  return element;
}

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

  it("builds a live-friendly summary with resolved and remaining counts", () => {
    const firstSuggestion = makeSuggestion({ id: "s-1" });
    const secondSuggestion = makeSuggestion({ id: "s-2" });

    expect(
      buildResultsSummary(
        [firstSuggestion, secondSuggestion],
        {
          successCount: 1,
          failedSuggestions: [
            {
              suggestion: secondSuggestion,
              reason: "not-found",
              message: "Anchor no encontrado",
            },
          ],
          pendingAfter: {
            pendingStylisticArtifacts: 1,
            hasPendingStylisticArtifacts: true,
            trackChangesActive: true,
          },
          documentState: "pending-review",
          trackChangesActivatedForBatch: true,
        },
        [],
        true
      )
    ).toBe(
      "Sobre selección — Te faltan 1 de 1 sugerencia aplicada por revisar. Todavía no resolviste ninguna. 1 no encontrada(s) en el texto."
    );
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
    expect(li.querySelector('[data-action="accept"]')?.textContent).toBe("Entendido");
    expect(li.querySelector('[data-action="reject"]')?.textContent).toBe("Ignorar");
    expect(li.querySelector(".result-type-badge--comment")?.textContent).toBe("comentario");
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
    const clickable = requireElement(li, ".card-clickable-area");

    expect(requireElement(li, ".result-original").textContent).toBe("fragmento exacto");
    clickable.click();

    expect(taskpaneMocks.navigateToText).toHaveBeenCalledWith(suggestion);
  });

  it('renders "No encontrado" cards after actionable suggestions on initial paint', async () => {
    const doc = createTaskpaneDocument();
    const firstSuggestion = makeSuggestion({ id: "s-1", anchor: "primero" });
    const missingSuggestion = makeSuggestion({ id: "s-missing", anchor: "faltante" });
    const secondSuggestion = makeSuggestion({ id: "s-2", anchor: "segundo" });

    const liItems = await renderViaEmitter(
      doc,
      [firstSuggestion, missingSuggestion, secondSuggestion],
      ["s-missing"]
    );

    expect(requireElement(liItems[0], ".result-original").textContent).toBe("primero");
    expect(requireElement(liItems[1], ".result-original").textContent).toBe("segundo");
    expect(requireElement(liItems[2], ".result-failed").textContent).toBe(
      'No encontrado: "faltante"'
    );
  });
});
