import type { ResultsPanelCardState } from "./ResultsPanelStore.types";
import type {
  ResultsPanelCardBucket,
  ResultsPanelChipCounts,
  ResultsPanelFilter,
} from "./ResultsPanelFilters.types";

/** Returns the chip bucket that a card contributes to. */
export function getResultsPanelCardBucket(card: ResultsPanelCardState): ResultsPanelCardBucket {
  if (card.isFailed || card.cardGroup === "not-found") {
    return "failed";
  }
  if (card.cardGroup === "processed") {
    if (card.state === "accepted") return "accepted";
    if (card.state === "rejected") return "rejected";
    return "failed";
  }
  return card.suggestion.severity;
}

/** Counts every chip bucket plus the total in a single pass. */
export function computeResultsPanelChipCounts(
  cards: readonly ResultsPanelCardState[]
): ResultsPanelChipCounts {
  const counts = {
    all: cards.length,
    high: 0,
    medium: 0,
    low: 0,
    accepted: 0,
    rejected: 0,
    failed: 0,
  };
  for (const card of cards) {
    counts[getResultsPanelCardBucket(card)] += 1;
  }
  return counts;
}

/** Returns cards visible under the current filter. */
export function selectResultsPanelVisibleCards(
  cards: readonly ResultsPanelCardState[],
  filter: ResultsPanelFilter
): readonly ResultsPanelCardState[] {
  if (filter === "all") {
    return cards;
  }
  return cards.filter((card) => getResultsPanelCardBucket(card) === filter);
}
