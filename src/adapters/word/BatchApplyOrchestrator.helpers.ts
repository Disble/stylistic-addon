import type { Suggestion } from "../../domain/suggestion/Suggestion.types";

/** Returns true when two hints belong to the same comparable snapshot scope. */
export function areComparableSnapshotHints(
  left: Suggestion["positionHint"],
  right: Suggestion["positionHint"]
): left is NonNullable<Suggestion["positionHint"]> {
  return (
    left?.source === "snapshot" &&
    right?.source === "snapshot" &&
    left.snapshotVersion === right.snapshotVersion &&
    left.paragraphId === right.paragraphId
  );
}
