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

/**
 * Orders suggestions from the last context in the document to the first, so
 * applying one never edits text inside a context still waiting to be found.
 * Contexts that end later go first; on a tie, the one that starts later. A
 * suggestion whose context is not in the text goes last, keeping its order.
 */
export function sortByContextPosition(
  suggestions: Suggestion[],
  documentText: string
): Suggestion[] {
  const text = normalizeLineBreaks(documentText);
  const located = suggestions.map((suggestion) => {
    const context = normalizeLineBreaks(suggestion.context);
    const start = text.indexOf(context);
    return { suggestion, start, end: start + context.length };
  });

  return located
    .sort((left, right) => {
      if (left.start < 0 || right.start < 0) {
        return Number(left.start < 0) - Number(right.start < 0);
      }

      return right.end - left.end || right.start - left.start;
    })
    .map(({ suggestion }) => suggestion);
}

/** Word separates paragraphs with `\r`; analyzed text may use `\n` or `\r\n`. */
function normalizeLineBreaks(value: string): string {
  return value.replace(/\r\n?/g, "\n");
}
