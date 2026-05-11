/** Semantic side of a replace tracked-change pair. */
export type ReplaceTrackedChangeSide = "Added" | "Deleted";

/** Compact tracked-change payload used in observability snapshots. */
export type TrackedChangeLogEntry = {
  type: string;
};
