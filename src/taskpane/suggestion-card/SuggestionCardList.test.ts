import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  createTaskpaneDocument,
  makeSuggestion,
  resetTaskpaneHarness,
  teardownTaskpaneHarness,
} from "../TaskpaneTestHelper";
import { createSuggestionCard } from "./SuggestionCardElements";
import { moveSuggestionCardToEnd } from "./SuggestionCardList";

describe("SuggestionCardList", () => {
  beforeEach(() => {
    resetTaskpaneHarness();
    globalThis.document = createTaskpaneDocument() as unknown as Document;
  });

  afterEach(() => {
    teardownTaskpaneHarness();
  });

  it("moves processed cards after active cards but before not-found cards", () => {
    const list = document.createElement("ul");
    const first = createSuggestionCard(makeSuggestion({ id: "s-1" }), []);
    const second = createSuggestionCard(makeSuggestion({ id: "s-2" }), []);
    const missingSuggestion = makeSuggestion({ id: "s-missing" });
    const missing = createSuggestionCard(missingSuggestion, [
      {
        suggestion: missingSuggestion,
        reason: "not-found",
        message: "Anchor no encontrado",
      },
    ]);

    list.appendChild(first.li);
    list.appendChild(second.li);
    list.appendChild(missing.li);

    moveSuggestionCardToEnd(first.li);

    expect(list.children[0]).toBe(second.li);
    expect(list.children[1]).toBe(first.li);
    expect(list.children[2]).toBe(missing.li);
    expect(first.li.dataset.cardGroup).toBe("processed");
  });
});
