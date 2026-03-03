/* global Word */

import { Suggestion, InsertionResult } from "./types";

/**
 * All Office.js / Word API interactions live here.
 * This module never contains business logic (SRP).
 */

export async function getDocumentText(): Promise<string> {
  return Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text;
  });
}

export async function insertSuggestionsAsTrackedChanges(
  suggestions: Suggestion[]
): Promise<InsertionResult> {
  return Word.run(async (context) => {
    // Preserve current tracking mode
    context.document.load("changeTrackingMode");
    await context.sync();
    const previousMode = context.document.changeTrackingMode;

    // Activate track changes
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();

    const failed: Suggestion[] = [];
    let successCount = 0;

    // Enqueue all search+replace operations before syncing
    const searchResults = suggestions.map((suggestion) => {
      const results = context.document.body.search(suggestion.originalText, {
        matchCase: true,
        matchWholeWord: false,
      });
      results.load("items");
      return { suggestion, results };
    });

    await context.sync();

    // Apply replacements for found matches
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

    // Restore previous tracking mode
    context.document.changeTrackingMode = previousMode;
    await context.sync();

    return { successCount, failedSuggestions: failed };
  });
}
