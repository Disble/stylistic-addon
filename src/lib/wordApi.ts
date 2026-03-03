/* global Word */

/**
 * Word API layer — all Office.js / Word interactions live here.
 *
 * This module is the only place that imports or references the `Word` global.
 * It exposes pure async functions that the task-pane orchestrator consumes.
 * Business logic (analysis, suggestion generation) MUST NOT live here (SRP).
 *
 * @module wordApi
 */

import { Suggestion, InsertionResult } from "./types";

/**
 * Reads and returns the full plain-text content of the active document.
 *
 * @returns The document body text. Returns an empty string for blank documents.
 */
export async function getDocumentText(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text;
  });
}

/**
 * Inserts an array of suggestions as native Word tracked changes.
 *
 * **Flow (preserve-and-restore pattern):**
 * 1. Reads and stores the current `changeTrackingMode`.
 * 2. Sets mode to `TrackAll` so every `insertText` is recorded as a revision.
 * 3. Batches all `body.search()` calls in a single `context.sync()` for performance.
 * 4. Replaces the target occurrence of each suggestion's `originalText`.
 * 5. Restores the previous tracking mode.
 *
 * The user can then accept/reject each change from Word's native Review tab.
 *
 * @param suggestions - Array of {@link Suggestion} objects to apply.
 * @returns An {@link InsertionResult} with success count and any suggestions
 *          whose `originalText` was not found in the document.
 */
export async function insertSuggestionsAsTrackedChanges(
  suggestions: Suggestion[]
): Promise<InsertionResult> {
  if (suggestions.length === 0) {
    return { successCount: 0, failedSuggestions: [] };
  }

  return Word.run(async (context) => {
    // 1. Preserve current tracking mode
    context.document.load("changeTrackingMode");
    await context.sync();
    const previousMode = context.document.changeTrackingMode;

    // 2. Activate track changes
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();

    const failed: Suggestion[] = [];
    let successCount = 0;

    try {
      // 3. Enqueue all search operations in a single batch (one sync, not one per suggestion)
      const searchResults = suggestions.map((suggestion) => {
        const results = context.document.body.search(suggestion.originalText, {
          matchCase: true,
          matchWholeWord: false,
        });
        results.load("items");
        return { suggestion, results };
      });

      await context.sync();

      // 4. Apply replacements for found matches
      for (const { suggestion, results } of searchResults) {
        if (results.items.length === 0) {
          failed.push(suggestion);
          continue;
        }

        const targetIndex =
          suggestion.paragraphIndex !== undefined &&
          suggestion.paragraphIndex < results.items.length
            ? suggestion.paragraphIndex
            : 0;

        results.items[targetIndex].insertText(
          suggestion.suggestedText,
          Word.InsertLocation.replace
        );
        successCount++;
      }

      await context.sync();
    } finally {
      // 5. ALWAYS restore previous tracking mode, even on error
      context.document.changeTrackingMode = previousMode;
      await context.sync();
    }

    return { successCount, failedSuggestions: failed };
  });
}
