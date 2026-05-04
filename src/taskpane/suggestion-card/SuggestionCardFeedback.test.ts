import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  createTaskpaneDocument,
  makeSuggestion,
  resetTaskpaneHarness,
  teardownTaskpaneHarness,
} from "../TaskpaneTestHelper";
import { createSuggestionCard } from "./SuggestionCardElements";
import {
  getSuggestionFeedbackComment,
  wireSuggestionFeedbackToggle,
} from "./SuggestionCardFeedback";

describe("SuggestionCardFeedback", () => {
  beforeEach(() => {
    resetTaskpaneHarness();
    globalThis.document = createTaskpaneDocument() as unknown as Document;
  });

  afterEach(() => {
    teardownTaskpaneHarness();
  });

  it("toggles the feedback accordion from the feedback button", () => {
    const card = createSuggestionCard(makeSuggestion(), []);
    const button = card.li.querySelector('[data-action="feedback"]') as HTMLButtonElement | null;
    const accordion = card.li.querySelector(".feedback-accordion");

    wireSuggestionFeedbackToggle(card.li);

    expect(accordion?.classList.contains("feedback-accordion--open")).toBe(false);
    button?.click();
    expect(accordion?.classList.contains("feedback-accordion--open")).toBe(true);
    button?.click();
    expect(accordion?.classList.contains("feedback-accordion--open")).toBe(false);
  });

  it("returns trimmed comments and omits empty feedback", () => {
    const card = createSuggestionCard(makeSuggestion(), []);
    const textarea = card.li.querySelector(".feedback-textarea") as HTMLTextAreaElement | null;

    if (!textarea) throw new Error("Missing feedback textarea");

    textarea.value = "   Muy buen cambio   ";
    expect(getSuggestionFeedbackComment(card.li)).toBe("Muy buen cambio");

    textarea.value = "   ";
    expect(getSuggestionFeedbackComment(card.li)).toBeUndefined();
  });
});
