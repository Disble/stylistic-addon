/** Returns the optional free-text feedback comment associated with a card. */
export function getSuggestionFeedbackComment(
  li: HTMLElement,
): string | undefined {
  const textarea = li.querySelector(".feedback-textarea") as
    | (HTMLTextAreaElement & { value?: string })
    | null;
  const commentText = textarea?.value?.trim();
  return commentText && commentText.length > 0 ? commentText : undefined;
}

/** Wires the feedback button to its accordion UI. */
export function wireSuggestionFeedbackToggle(li: HTMLElement): void {
  const feedbackBtnEl = li.querySelector(
    '[data-action="feedback"]',
  ) as HTMLButtonElement | null;
  const accordionEl = li.querySelector(
    ".feedback-accordion",
  ) as HTMLElement | null;

  if (feedbackBtnEl && accordionEl) {
    feedbackBtnEl.addEventListener("click", () => {
      accordionEl.classList.toggle("feedback-accordion--open");
    });
  }
}
