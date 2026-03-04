/* global Word */

/**
 * Word API layer — all Office.js / Word interactions live here.
 *
 * This module is the only place that imports or references the `Word` global.
 * It exposes pure async functions that the task-pane orchestrator consumes.
 * Business logic and backend communication MUST NOT live here (SRP).
 *
 * Key design decisions:
 * - **Strategy pattern for tracked changes:** suggestions are classified as
 *   insert, delete, or replace. All types use the normal Word API under
 *   `TrackAll` mode via `insertText(replace)` / `range.delete()`. Word's
 *   native tracking engine handles revision creation.
 * - **Preserve-and-restore pattern:** the document's `changeTrackingMode` is
 *   saved before modification and restored in a `finally` block, even on error.
 * - **Batched application:** suggestions are applied in groups of
 *   {@link WORD_API_BATCH_SIZE} via separate `Word.run` calls. Each batch
 *   is an independent commit — if batch N fails, batches 1..(N-1) are already
 *   persisted in the document. This maximizes reliability for large documents.
 * - **Progress callbacks:** the caller receives updates after each batch,
 *   enabling real-time UI progress reporting.
 *
 * @module wordApi
 */

import {
  Suggestion,
  InsertionResult,
  ProgressCallback,
  ChangeType,
} from "./types";
import { WORD_API_BATCH_SIZE } from "./config";

// ---------------------------------------------------------------------------
// Document Reading
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Tracked Changes Application
// ---------------------------------------------------------------------------

/**
 * Applies an array of suggestions as native Word tracked changes, processing
 * them in batches for reliability and progress reporting.
 *
 * Uses the **Strategy pattern** internally:
 * - Delete suggestions: `range.delete()` under TrackAll.
 * - Insert/Replace suggestions: `insertText(replace)` under TrackAll.
 *
 * Word's native tracking engine creates the revision marks. For replacements,
 * `insertText(replace)` under `TrackAll` lets Word handle the replacement
 * as a single atomic operation.
 *
 * Each `Word.run` is independent — changes are committed after each batch.
 * If batch 5 fails, batches 1–4 are already persisted.
 *
 * @param suggestions - Array of {@link Suggestion} objects to apply.
 * @param onProgress  - Optional callback invoked after each batch completes.
 * @returns An {@link InsertionResult} with success count and failed suggestions.
 */
export async function applySuggestionsInBatches(
  suggestions: Suggestion[],
  onProgress?: ProgressCallback
): Promise<InsertionResult> {
  if (suggestions.length === 0) {
    return { successCount: 0, failedSuggestions: [] };
  }

  const batches = createBatches(suggestions, WORD_API_BATCH_SIZE);
  const allFailed: Suggestion[] = [];
  let totalSuccess = 0;

  // Phase 1: Save tracking mode and enable TrackAll
  const previousMode = await saveAndEnableTrackAll();

  try {
    // Phase 2: Apply each batch in a separate Word.run
    for (const batch of batches) {
      const batchResult = await applyBatch(batch);

      totalSuccess += batchResult.successCount;
      allFailed.push(...batchResult.failedSuggestions);

      if (onProgress) {
        onProgress(
          "applying",
          totalSuccess,
          suggestions.length,
          `Aplicando sugerencia ${totalSuccess} de ${suggestions.length}...`
        );
      }
    }
  } finally {
    // Phase 3: Always restore the previous tracking mode
    await restoreTrackingMode(previousMode);
  }

  return { successCount: totalSuccess, failedSuggestions: allFailed };
}

// ---------------------------------------------------------------------------
// Internal: Change Type Classification (Strategy Pattern)
// ---------------------------------------------------------------------------

/**
 * Determines the type of tracked change operation for a suggestion.
 *
 * - `"delete"`: originalText is non-empty, suggestedText is empty.
 * - `"insert"`: originalText is empty, suggestedText is non-empty.
 * - `"replace"`: both are non-empty and different.
 *
 * @param suggestion - The suggestion to classify.
 * @returns The {@link ChangeType} for this suggestion.
 */
function classifyChange(suggestion: Suggestion): ChangeType {
  const hasOriginal = suggestion.originalText.length > 0;
  const hasSuggested = suggestion.suggestedText.length > 0;

  if (hasOriginal && !hasSuggested) return "delete";
  if (!hasOriginal && hasSuggested) return "insert";
  return "replace";
}

// ---------------------------------------------------------------------------
// Internal: Tracking Mode Management
// ---------------------------------------------------------------------------

/**
 * Saves the document's current `changeTrackingMode` and sets it to `TrackAll`.
 *
 * @returns The previous tracking mode value, to be restored later.
 */
async function saveAndEnableTrackAll(): Promise<string> {
  return Word.run(async (context) => {
    context.document.load("changeTrackingMode");
    await context.sync();

    const previousMode = context.document.changeTrackingMode;
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();

    return previousMode;
  });
}

/**
 * Restores the document's `changeTrackingMode` to a previously saved value.
 * Swallows errors to avoid masking the original error in a `finally` block.
 *
 * @param mode - The tracking mode to restore.
 */
async function restoreTrackingMode(mode: string): Promise<void> {
  try {
    await Word.run(async (context) => {
      context.document.changeTrackingMode =
        mode as Word.ChangeTrackingMode;
      await context.sync();
    });
  } catch {
    // Swallow — restoring the mode is best-effort.
    // Failing here should not mask the primary operation's result.
  }
}

// ---------------------------------------------------------------------------
// Internal: Batch Application (Strategy Dispatcher)
// ---------------------------------------------------------------------------

/**
 * Applies a single batch of suggestions using the Strategy pattern.
 *
 * All suggestions are applied in a single `Word.run` with TrackAll active.
 * Each suggestion is dispatched to the appropriate strategy based on its type:
 * - Delete: `range.delete()`
 * - Insert/Replace: `range.insertText(suggestedText, replace)`
 *
 * @param batch - Array of suggestions to apply in this batch.
 * @returns An {@link InsertionResult} for this batch only.
 */
async function applyBatch(batch: Suggestion[]): Promise<InsertionResult> {
  return Word.run(async (context) => {
    const failed: Suggestion[] = [];
    let successCount = 0;

    // Enqueue all searches in a single batch (no syncs yet)
    const searchResults = batch.map((suggestion) => {
      const results = context.document.body.search(
        suggestion.originalText,
        { matchCase: true, matchWholeWord: false }
      );
      results.load("items");
      return { suggestion, results };
    });

    // Execute all searches in one round-trip
    await context.sync();

    // Apply changes for found matches using the appropriate strategy
    for (const { suggestion, results } of searchResults) {
      if (results.items.length === 0) {
        failed.push(suggestion);
        continue;
      }

      const type = classifyChange(suggestion);
      if (type === "delete") {
        results.items[0].delete();
      } else {
        // "insert" or "replace" — Word handles both via insertText
        results.items[0].insertText(
          suggestion.suggestedText,
          Word.InsertLocation.replace
        );
      }
      successCount++;
    }

    // Commit all changes in one round-trip
    await context.sync();

    return { successCount, failedSuggestions: failed };
  });
}

// ---------------------------------------------------------------------------
// PoC: Tracked Change Test Methods
// ---------------------------------------------------------------------------

/**
 * PoC Method A: Uses `insertText(replace)` under `TrackAll` mode.
 *
 * This is the standard Word API approach. Word's native tracking engine
 * handles revision creation. According to the documentation, this should
 * produce a single combined tracked change for replacements.
 *
 * @param originalText  - Exact text to find in the document (case-sensitive).
 * @param suggestedText - Replacement text for the tracked change.
 * @returns An object with `success` status and a descriptive `message`.
 */
export async function pocInsertTextReplace(
  originalText: string,
  suggestedText: string
): Promise<{ success: boolean; message: string }> {
  return Word.run(async (context) => {
    // Save and enable TrackAll
    context.document.load("changeTrackingMode");
    await context.sync();
    const previousMode = context.document.changeTrackingMode;
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await context.sync();

    try {
      const results = context.document.body.search(originalText, {
        matchCase: true,
        matchWholeWord: false,
      });
      results.load("items");
      await context.sync();

      if (results.items.length === 0) {
        return {
          success: false,
          message: `"${originalText}" no encontrado en el documento.`,
        };
      }

      if (!suggestedText) {
        // Delete only
        results.items[0].delete();
      } else {
        // Replace
        results.items[0].insertText(
          suggestedText,
          Word.InsertLocation.replace
        );
      }
      await context.sync();

      const action = suggestedText
        ? `"${originalText}" → "${suggestedText}"`
        : `"${originalText}" eliminado`;

      return {
        success: true,
        message: `[insertText] ${action} — aplicado con TrackAll.`,
      };
    } finally {
      context.document.changeTrackingMode =
        previousMode as Word.ChangeTrackingMode;
      await context.sync();
    }
  });
}

/**
 * PoC Method B: Uses OOXML `<w:del>` + `<w:ins>` markup with tracking OFF.
 *
 * Builds a flat OPC OOXML package containing tracked change markup and
 * inserts it via `insertOoxml()`. According to research, Word may
 * strip/reprocess revision markup during `insertOoxml()`.
 *
 * @param originalText  - Exact text to find in the document (case-sensitive).
 * @param suggestedText - Replacement text for the tracked change.
 * @returns An object with `success` status and a descriptive `message`.
 */
export async function pocInsertOoxmlReplace(
  originalText: string,
  suggestedText: string
): Promise<{ success: boolean; message: string }> {
  if (!suggestedText) {
    // For deletion, delegate to insertText approach (OOXML not needed)
    return pocInsertTextReplace(originalText, suggestedText);
  }

  return Word.run(async (context) => {
    // Save and disable tracking so the OOXML insertion isn't double-tracked
    context.document.load("changeTrackingMode");
    await context.sync();
    const previousMode = context.document.changeTrackingMode;
    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    await context.sync();

    try {
      const results = context.document.body.search(originalText, {
        matchCase: true,
        matchWholeWord: false,
      });
      results.load("items");
      await context.sync();

      if (results.items.length === 0) {
        return {
          success: false,
          message: `"${originalText}" no encontrado en el documento.`,
        };
      }

      const range = results.items[0];

      // Extract formatting from the original range
      const rangeOoxml = range.getOoxml();
      await context.sync();
      const runProps = extractRunProperties(rangeOoxml.value);

      // Build and insert the tracked change OOXML
      const ooxml = buildTrackedChangeOoxml(
        originalText,
        suggestedText,
        runProps
      );
      range.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();

      return {
        success: true,
        message: `[OOXML] "${originalText}" → "${suggestedText}" — insertado con OOXML markup.`,
      };
    } finally {
      context.document.changeTrackingMode =
        previousMode as Word.ChangeTrackingMode;
      await context.sync();
    }
  });
}

// ---------------------------------------------------------------------------
// Internal: OOXML Building
// ---------------------------------------------------------------------------

/**
 * Builds an Office Open XML (flat OPC) package that represents a single
 * tracked change: a deletion of `original` and an insertion of `replacement`,
 * both under the same author and timestamp.
 *
 * Word renders adjacent `<w:del>` + `<w:ins>` elements with the same author
 * and date as one combined revision in the Review pane.
 *
 * The package includes the required `/_rels/.rels` relationship part to
 * ensure Word can locate the document content.
 *
 * @param original    - The text being replaced (shown as strikethrough).
 * @param replacement - The new text (shown as underlined insertion).
 * @param runPropertiesXml - Optional serialized `<w:rPr>` element to preserve
 *                           formatting from the original range.
 * @returns A flat OPC XML string suitable for `Range.insertOoxml()`.
 */
function buildTrackedChangeOoxml(
  original: string,
  replacement: string,
  runPropertiesXml?: string | null
): string {
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  const escapedOriginal = escapeXml(original);
  const escapedReplacement = escapeXml(replacement);

  // Build run properties element if available
  const rPrBlock = runPropertiesXml
    ? `                ${runPropertiesXml}\n`
    : "";

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">',
    '  <pkg:part pkg:name="/_rels/.rels"',
    '    pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">',
    "    <pkg:xmlData>",
    '      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '        <Relationship Id="rId1"',
    '          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"',
    '          Target="word/document.xml"/>',
    "      </Relationships>",
    "    </pkg:xmlData>",
    "  </pkg:part>",
    '  <pkg:part pkg:name="/word/document.xml"',
    '    pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">',
    "    <pkg:xmlData>",
    '      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    "        <w:body>",
    "          <w:p>",
    `            <w:del w:id="0" w:author="Stylistic" w:date="${now}">`,
    "              <w:r>",
    `${rPrBlock}                <w:delText xml:space="preserve">${escapedOriginal}</w:delText>`,
    "              </w:r>",
    "            </w:del>",
    `            <w:ins w:id="1" w:author="Stylistic" w:date="${now}">`,
    "              <w:r>",
    `${rPrBlock}                <w:t xml:space="preserve">${escapedReplacement}</w:t>`,
    "              </w:r>",
    "            </w:ins>",
    "          </w:p>",
    "        </w:body>",
    "      </w:document>",
    "    </pkg:xmlData>",
    "  </pkg:part>",
    "</pkg:package>",
  ].join("\n");
}

// ---------------------------------------------------------------------------
// Internal: OOXML Parsing
// ---------------------------------------------------------------------------

/**
 * Extracts the first `<w:rPr>` (run properties) element from a flat OPC
 * OOXML string. This preserves the original formatting when building
 * tracked change OOXML.
 *
 * @param ooxml - The flat OPC OOXML string from `Range.getOoxml()`.
 * @returns The serialized `<w:rPr>` XML string, or `null` if not found.
 */
function extractRunProperties(ooxml: string): string | null {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(ooxml, "application/xml");
    const nsW =
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    const rPr = doc.getElementsByTagNameNS(nsW, "rPr")[0];
    if (!rPr) return null;

    return new XMLSerializer().serializeToString(rPr);
  } catch {
    // If parsing fails, fall back to no formatting preservation
    return null;
  }
}

// ---------------------------------------------------------------------------
// Internal: XML Escaping
// ---------------------------------------------------------------------------

/**
 * Escapes special XML characters to prevent injection in OOXML content.
 *
 * @param str - Raw string that may contain XML-special characters.
 * @returns The XML-safe escaped string.
 */
function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// ---------------------------------------------------------------------------
// Internal: Utility
// ---------------------------------------------------------------------------

/**
 * Splits an array into fixed-size sub-arrays (batches).
 *
 * @param items     - The array to split.
 * @param batchSize - Maximum items per batch.
 * @returns An array of batches.
 */
function createBatches<T>(items: T[], batchSize: number): T[][] {
  const batches: T[][] = [];
  for (let i = 0; i < items.length; i += batchSize) {
    batches.push(items.slice(i, i + batchSize));
  }
  return batches;
}
