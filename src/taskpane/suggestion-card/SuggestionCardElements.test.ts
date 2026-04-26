import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { SuggestionApplicationFailure } from "../../domain/DocumentApplication.types";
import {
  createTaskpaneDocument,
  makeSuggestion,
  resetTaskpaneHarness,
  teardownTaskpaneHarness,
} from "../TaskpaneTestHelper";
import { createSuggestionCard } from "./SuggestionCardElements";

describe("SuggestionCardElements", () => {
  beforeEach(() => {
    resetTaskpaneHarness();
    globalThis.document = createTaskpaneDocument() as unknown as Document;
  });

  afterEach(() => {
    teardownTaskpaneHarness();
  });

  it("renders comment-only suggestions without diff blocks and with text actions", () => {
    const suggestion = makeSuggestion({
      id: "s-comment",
      type: "comment-only",
      suggestedText: undefined,
    });

    const card = createSuggestionCard(suggestion, []);

    expect(card.li.querySelector(".card-diff")).toBeNull();
    expect(card.li.querySelector('[data-action="accept"]')?.textContent).toBe(
      "Entendido",
    );
    expect(card.li.querySelector('[data-action="reject"]')?.textContent).toBe(
      "Ignorar",
    );
    expect(
      card.li.querySelector(".result-type-badge--comment")?.textContent,
    ).toBe("comentario");
  });

  it("renders failed suggestions without feedback or action controls", () => {
    const suggestion = makeSuggestion({ id: "s-fail", anchor: "faltante" });
    const failure: SuggestionApplicationFailure = {
      suggestion,
      reason: "not-found",
      message: "Anchor no encontrado",
    };

    const card = createSuggestionCard(suggestion, [failure]);

    expect(card.isFailed).toBe(true);
    expect(card.isNotFoundFailure).toBe(true);
    expect(card.li.dataset.cardGroup).toBe("not-found");
    expect(card.li.querySelector(".result-failed")?.textContent).toBe(
      'No encontrado: "faltante"',
    );
    expect(card.li.querySelector('[data-action="feedback"]')).toBeNull();
    expect(card.li.querySelector(".feedback-accordion")).toBeNull();
  });
});
