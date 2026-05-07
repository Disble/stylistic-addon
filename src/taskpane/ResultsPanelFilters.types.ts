/** Mutually exclusive chip buckets a card can belong to. */
export type ResultsPanelCardBucket = "high" | "medium" | "low" | "accepted" | "rejected" | "failed";

/** Filter selected via the chips toolbar; "all" means no filter applied. */
export type ResultsPanelFilter = "all" | ResultsPanelCardBucket;

/** Per-bucket card counts plus the "all" total. */
export type ResultsPanelChipCounts = Readonly<{
  all: number;
  high: number;
  medium: number;
  low: number;
  accepted: number;
  rejected: number;
  failed: number;
}>;
