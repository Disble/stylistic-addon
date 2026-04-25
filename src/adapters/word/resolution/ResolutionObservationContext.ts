import type { ReplaceTrackedChangeSide } from "./ReplaceResolutionStrategyContext";

/** Compact tracked-change payload used in observability snapshots. */
export type TrackedChangeLogEntry = {
  type: string;
};

/** Builds one stable tracked-change entry for cross-phase workflow logs. */
export function describeTrackedChangeForLog(
  trackedChange: Word.TrackedChange,
): TrackedChangeLogEntry {
  return {
    type: trackedChange.type ?? "unknown",
  };
}

/** Builds a compact tracked-change list so one workflow attempt can be reconstructed later. */
export function describeTrackedChangesForLog(
  trackedChanges: Word.TrackedChange[],
): TrackedChangeLogEntry[] {
  return trackedChanges.map((trackedChange) =>
    describeTrackedChangeForLog(trackedChange),
  );
}

/** Builds a compact comma-separated type summary for one tracked-change collection. */
export function formatTrackedChangeTypesForLog(
  trackedChanges: Word.TrackedChange[],
): string {
  return trackedChanges
    .map((trackedChange) => trackedChange.type ?? "unknown")
    .join(",");
}

/** Resolves one stale preferred CC to its fresh logical equivalent from the current locate pass. */
export function resolveFreshPreferredCandidate(
  candidates: Word.ContentControl[],
  preferredCc?: Word.ContentControl,
): Word.ContentControl | null {
  if (!preferredCc) {
    return null;
  }

  const preferredTag = preferredCc.tag;
  const preferredTitle = preferredCc.title ?? "";

  return (
    candidates.find(
      (candidate) =>
        candidate.tag === preferredTag &&
        (candidate.title ?? "") === preferredTitle,
    ) ?? null
  );
}

/** Keeps the fresh logical successor first without ever reusing the old proxy object. */
export function prioritizeFreshPreferredCandidate(
  candidates: Word.ContentControl[],
  preferredCc: Word.ContentControl | null,
): Word.ContentControl[] {
  if (!preferredCc) {
    return candidates;
  }

  return [
    preferredCc,
    ...candidates.filter((candidate) => candidate !== preferredCc),
  ];
}

/** Returns the first tracked change for the requested semantic side. */
export function findTrackedChangeByType(
  trackedChanges: Word.TrackedChange[],
  trackedChangeType: ReplaceTrackedChangeSide,
): Word.TrackedChange | null {
  return (
    trackedChanges.find(
      (trackedChange) => trackedChange.type === trackedChangeType,
    ) ?? null
  );
}
