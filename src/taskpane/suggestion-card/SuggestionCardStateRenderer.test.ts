import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  createTaskpaneDocument,
  makeSuggestion,
  resetTaskpaneHarness,
  teardownTaskpaneHarness,
} from "../TaskpaneTestHelper";
import { createSuggestionCard } from "./SuggestionCardElements";
import { applySuggestionCardState } from "./SuggestionCardStateRenderer";

describe("SuggestionCardStateRenderer", () => {
  beforeEach(() => {
    resetTaskpaneHarness();
    globalThis.document = createTaskpaneDocument() as unknown as Document;
  });

  afterEach(() => {
    teardownTaskpaneHarness();
  });

  it("marks accepted cards as terminal and removes actions", () => {
    const card = createSuggestionCard(makeSuggestion(), []);

    applySuggestionCardState(card.li, "accepted", null, null);

    expect(card.li.classList.contains("result-accepted")).toBe(true);
    expect(card.li.querySelector(".result-actions")).toBeNull();
  });

  it("renders manual-review states with a note and status message", () => {
    const card = createSuggestionCard(makeSuggestion(), []);
    const statusBar = document.getElementById("status-bar");

    applySuggestionCardState(
      card.li,
      "ambiguous-location",
      null,
      null,
      "La ubicación de la sugerencia es ambigua."
    );

    expect(card.li.classList.contains("result-ambiguous-location")).toBe(true);
    expect(card.li.querySelector(".result-ambiguous-location-note")?.textContent).toBe(
      "(resolución ambigua; reanalizá la sugerencia)"
    );
    expect(statusBar?.textContent).toBe("La ubicación de la sugerencia es ambigua.");
  });
});
