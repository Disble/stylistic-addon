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

/** Resolves one stale preferred CC to its fresh logical equivalent from the current locate pass. */
export function resolveFreshPreferredCandidate(
  rankedCandidates: Word.ContentControl[],
  preferredCc?: Word.ContentControl,
): Word.ContentControl | null {
  if (!preferredCc) {
    return null;
  }

  const preferredTag = preferredCc.tag;
  const preferredTitle = preferredCc.title ?? "";

  return (
    rankedCandidates.find(
      (candidate) =>
        candidate.tag === preferredTag &&
        (candidate.title ?? "") === preferredTitle,
    ) ?? null
  );
}

/** Keeps the fresh logical successor first without ever reusing the old proxy object. */
export function prioritizeFreshPreferredCandidate(
  rankedCandidates: Word.ContentControl[],
  preferredCc: Word.ContentControl | null,
): Word.ContentControl[] {
  if (!preferredCc) {
    return rankedCandidates;
  }

  return [
    preferredCc,
    ...rankedCandidates.filter((candidate) => candidate !== preferredCc),
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
