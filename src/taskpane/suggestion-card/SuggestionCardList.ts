import { SUGGESTION_CARD_REORDER_ANIMATION_MS } from "../../infrastructure/config";

/**
 * Runs a FLIP-style animation for suggestion-list reordering when the host DOM
 * provides layout APIs. Falls back to an immediate reorder in test/fake DOM.
 */
export function animateSuggestionListReorder(
  parent: HTMLElement,
  reorder: () => void,
): void {
  const cardsBefore = Array.from(parent.children) as HTMLElement[];
  const canAnimate =
    typeof globalThis.requestAnimationFrame === "function" &&
    cardsBefore.every(
      (card) => typeof card.getBoundingClientRect === "function",
    );

  if (!canAnimate) {
    reorder();
    return;
  }

  const firstRects = new Map(
    cardsBefore.map((card) => [card, card.getBoundingClientRect()]),
  );

  reorder();

  const cardsAfter = Array.from(parent.children) as HTMLElement[];
  for (const card of cardsAfter) {
    const firstRect = firstRects.get(card);
    if (!firstRect) {
      continue;
    }

    const lastRect = card.getBoundingClientRect();
    const deltaY = firstRect.top - lastRect.top;
    if (deltaY === 0) {
      continue;
    }

    card.style.transition = "none";
    card.style.transform = `translateY(${deltaY}px)`;
  }

  globalThis.requestAnimationFrame(() => {
    for (const card of cardsAfter) {
      const firstRect = firstRects.get(card);
      if (!firstRect) {
        continue;
      }

      const lastRect = card.getBoundingClientRect();
      const deltaY = firstRect.top - lastRect.top;
      if (deltaY === 0) {
        continue;
      }

      card.style.transition = `transform ${SUGGESTION_CARD_REORDER_ANIMATION_MS}ms ease`;
      card.style.transform = "";
    }

    globalThis.setTimeout(() => {
      for (const card of cardsAfter) {
        card.style.transition = "";
        card.style.transform = "";
      }
    }, SUGGESTION_CARD_REORDER_ANIMATION_MS);
  });
}

/** Returns the first card that represents a not-found failure. */
export function getFirstNotFoundCard(parent: HTMLElement): HTMLElement | null {
  return (
    (Array.from(parent.children) as HTMLElement[]).find(
      (card) => card.dataset.cardGroup === "not-found",
    ) ?? null
  );
}

/** Moves a terminally processed suggestion card after active cards. */
export function moveSuggestionCardToEnd(li: HTMLElement): void {
  const parent = li.parentElement;
  if (!parent) {
    return;
  }

  li.dataset.cardGroup = "processed";

  animateSuggestionListReorder(parent, () => {
    const firstNotFoundCard = getFirstNotFoundCard(parent);
    if (firstNotFoundCard && firstNotFoundCard !== li) {
      firstNotFoundCard.before(li);
      return;
    }

    parent.appendChild(li);
  });
}
