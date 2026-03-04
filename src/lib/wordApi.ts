/* global Word */

/**
 * Word API layer — all Office.js / Word interactions live here.
 *
 * This module is the only place that imports or references the `Word` global.
 * It exposes pure async functions that the task-pane orchestrator consumes.
 * Business logic and backend communication MUST NOT live here (SRP).
 *
 * Key design decisions:
 * - **OOXML as primary strategy:** all change types (insert, delete, replace)
 *   are applied via flat OPC OOXML packages containing tracked-change markup
 *   (`<w:del>`, `<w:ins>`) and a formatted Word comment with the justification.
 *   The comment shows as a margin balloon with bold category and justification
 *   text. The tracked change blue card shows "Stylistic" as author.
 * - **Per-suggestion processing:** each suggestion is applied in its own
 *   `Word.run` to avoid stale ranges after OOXML insertions shift text.
 * - **Preserve-and-restore pattern:** the document's `changeTrackingMode` is
 *   saved before modification and restored in a `finally` block, even on error.
 * - **Progress callbacks:** the caller receives updates after each suggestion,
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
 * Applies an array of suggestions as native Word tracked changes with
 * embedded justification comments, processing them one at a time via OOXML.
 *
 * Each suggestion is applied in its own `Word.run`:
 * 1. Search for the original text.
 * 2. Extract formatting (`<w:rPr>`) from the matched range.
 * 3. Disable tracking (to avoid double-tracking the OOXML insertion).
 * 4. Build and insert the OOXML package (tracked change + comment).
 * 5. Restore tracking mode.
 *
 * @param suggestions - Array of {@link Suggestion} objects to apply.
 * @param onProgress  - Optional callback invoked after each suggestion.
 * @returns An {@link InsertionResult} with success count and failed suggestions.
 */
export async function applySuggestionsInBatches(
  suggestions: Suggestion[],
  onProgress?: ProgressCallback
): Promise<InsertionResult> {
  if (suggestions.length === 0) {
    return { successCount: 0, failedSuggestions: [] };
  }

  const allFailed: Suggestion[] = [];
  let totalSuccess = 0;

  for (const suggestion of suggestions) {
    const result = await applySingleSuggestion(suggestion);

    if (result.success) {
      totalSuccess++;
    } else {
      allFailed.push(suggestion);
    }

    if (onProgress) {
      onProgress(
        "applying",
        totalSuccess + allFailed.length,
        suggestions.length,
        `Aplicando sugerencia ${totalSuccess + allFailed.length} de ${suggestions.length}...`
      );
    }
  }

  return { successCount: totalSuccess, failedSuggestions: allFailed };
}

// ---------------------------------------------------------------------------
// Internal: Single Suggestion Application
// ---------------------------------------------------------------------------

/**
 * Applies a single suggestion as a tracked change with an embedded comment.
 *
 * @param suggestion - The suggestion to apply.
 * @returns An object indicating success or failure.
 */
async function applySingleSuggestion(
  suggestion: Suggestion
): Promise<{ success: boolean }> {
  return Word.run(async (context) => {
    // Search for the original text
    const results = context.document.body.search(
      suggestion.originalText,
      { matchCase: true, matchWholeWord: false }
    );
    results.load("items");
    await context.sync();

    if (results.items.length === 0) {
      return { success: false };
    }

    const range = results.items[0];

    // Extract formatting from the original range
    const rangeOoxml = range.getOoxml();
    await context.sync();
    const runProps = extractRunProperties(rangeOoxml.value);

    // Save tracking mode and disable it (OOXML contains its own markup)
    context.document.load("changeTrackingMode");
    await context.sync();
    const previousMode = context.document.changeTrackingMode;
    context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
    await context.sync();

    try {
      // Build and insert the OOXML package
      const type = classifyChange(suggestion);
      const ooxml = buildTrackedChangeOoxml(
        suggestion.originalText,
        suggestion.suggestedText,
        suggestion.justification,
        suggestion.category,
        type,
        runProps
      );
      range.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();

      return { success: true };
    } finally {
      // Restore tracking mode
      context.document.changeTrackingMode =
        previousMode as Word.ChangeTrackingMode;
      await context.sync();
    }
  });
}

// ---------------------------------------------------------------------------
// Internal: Change Type Classification (Strategy Pattern)
// ---------------------------------------------------------------------------

/**
 * Determines the type of tracked change operation for a suggestion.
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
// Comment Cleanup
// ---------------------------------------------------------------------------

/**
 * Deletes Stylistic comments whose tracked changes have been resolved
 * (accepted or rejected by the user). Keeps comments for still-pending
 * tracked changes and never touches comments from other authors.
 *
 * Pattern: Range Colocation
 * Uses the Word API's spatial range comparison to determine which comments
 * are still anchored to a pending tracked change. No in-memory registry
 * needed — the document itself is the source of truth.
 *
 * 1. Load all Stylistic comments and tracked changes.
 * 2. Get the document range for each (comment.getRange(), tc.getRange()).
 * 3. For each comment, check if any TC range overlaps with it.
 * 4. No overlap → orphaned (TC was resolved) → delete.
 * 5. Overlap found → TC still pending → keep.
 *
 * @returns Counts of deleted and kept Stylistic comments.
 */
export async function cleanupResolvedComments(): Promise<{
  deleted: number;
  kept: number;
}> {
  return Word.run(async (context) => {
    // Sync 1: load collections with properties
    const tracked = context.document.body.getTrackedChanges();
    const comments = context.document.body.getComments();
    tracked.load({ select: "author,type" });
    comments.load({ select: "authorName" });
    await context.sync();

    const stylisticComments = comments.items.filter(
      (c) => c.authorName === "Stylistic"
    );
    const stylisticTCs = tracked.items.filter(
      (tc) => tc.author === "Stylistic"
    );

    if (stylisticComments.length === 0) {
      return { deleted: 0, kept: 0 };
    }

    // All TCs resolved → delete all Stylistic comments, skip range comparison
    if (stylisticTCs.length === 0) {
      for (const comment of stylisticComments) {
        comment.delete();
      }
      await context.sync();
      return { deleted: stylisticComments.length, kept: 0 };
    }

    // Sync 2: get document ranges for each comment and TC
    const commentRanges = stylisticComments.map((c) => c.getRange());
    const tcRanges = stylisticTCs.map((tc) => tc.getRange());
    await context.sync();

    // Sync 3: compare each comment range against each TC range
    const comparisons: OfficeExtension.ClientResult<Word.LocationRelation>[][] =
      [];
    for (let i = 0; i < commentRanges.length; i++) {
      comparisons[i] = [];
      for (let j = 0; j < tcRanges.length; j++) {
        comparisons[i][j] = commentRanges[i].compareLocationWith(tcRanges[j]);
      }
    }
    await context.sync();

    // Evaluate: a comment is colocated with a TC if their ranges overlap
    const overlapping: Word.LocationRelation[] = [
      "Equal" as Word.LocationRelation,
      "Contains" as Word.LocationRelation,
      "ContainsStart" as Word.LocationRelation,
      "ContainsEnd" as Word.LocationRelation,
      "Inside" as Word.LocationRelation,
      "InsideStart" as Word.LocationRelation,
      "InsideEnd" as Word.LocationRelation,
      "OverlapsBefore" as Word.LocationRelation,
      "OverlapsAfter" as Word.LocationRelation,
    ];

    let deleted = 0;
    let kept = 0;

    for (let i = 0; i < stylisticComments.length; i++) {
      let hasColocatedTC = false;
      for (let j = 0; j < tcRanges.length; j++) {
        if (overlapping.includes(comparisons[i][j].value)) {
          hasColocatedTC = true;
          break;
        }
      }

      if (hasColocatedTC) {
        kept++;
      } else {
        stylisticComments[i].delete();
        deleted++;
      }
    }

    // Sync 4: execute deletes
    await context.sync();
    return { deleted, kept };
  });
}

// ---------------------------------------------------------------------------
// Internal: OOXML Building
// ---------------------------------------------------------------------------

/**
 * Builds a flat OPC OOXML package with tracked change markup and a
 * formatted Word comment containing the justification.
 *
 * Package structure (4 parts, no [Content_Types].xml):
 * - `/_rels/.rels` → points to word/document.xml
 * - `/word/_rels/document.xml.rels` → points to word/comments.xml
 * - `/word/document.xml` → tracked change + comment anchors
 * - `/word/comments.xml` → formatted justification (bold category + text)
 *
 * The comment appears as a margin balloon (when "Show Revisions in
 * Balloons" is active) with the category in bold and the justification
 * as a separate paragraph. The tracked change blue card shows
 * `w:author="Stylistic"` cleanly.
 *
 * @param original           - The text being replaced/deleted.
 * @param replacement        - The new text (empty for delete-only).
 * @param justification      - Reason for the change (comment body).
 * @param category           - Category label (bold in comment).
 * @param changeType         - The type of change (insert/delete/replace).
 * @param runPropertiesXml   - Optional `<w:rPr>` XML to preserve formatting.
 * @returns A flat OPC XML string suitable for `Range.insertOoxml()`.
 */
function buildTrackedChangeOoxml(
  original: string,
  replacement: string,
  justification: string,
  category: string,
  changeType: ChangeType,
  runPropertiesXml?: string | null
): string {
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  const escapedOriginal = escapeXml(original);
  const escapedReplacement = escapeXml(replacement);
  const escapedCategory = escapeXml(category);
  const escapedJustification = escapeXml(justification);

  const rPr = runPropertiesXml
    ? `                ${runPropertiesXml}\n`
    : "";

  // Build tracked change body based on type
  let changeBody = "";

  if (changeType === "delete" || changeType === "replace") {
    changeBody +=
      `            <w:del w:id="1" w:author="Stylistic" w:date="${now}">\n` +
      `              <w:r>\n` +
      `${rPr}                <w:delText xml:space="preserve">${escapedOriginal}</w:delText>\n` +
      `              </w:r>\n` +
      `            </w:del>\n`;
  }

  if (changeType === "insert" || changeType === "replace") {
    changeBody +=
      `            <w:ins w:id="2" w:author="Stylistic" w:date="${now}">\n` +
      `              <w:r>\n` +
      `${rPr}                <w:t xml:space="preserve">${escapedReplacement}</w:t>\n` +
      `              </w:r>\n` +
      `            </w:ins>\n`;
  }

  // Build comment paragraphs: bold category + each justification line as <w:p>
  const justificationParagraphs = justification
    .split("\n")
    .filter((line) => line.trim().length > 0)
    .map(
      (line) =>
        `          <w:p>\n` +
        `            <w:r>\n` +
        `              <w:t>${escapeXml(line)}</w:t>\n` +
        `            </w:r>\n` +
        `          </w:p>`
    )
    .join("\n");

  const categoryParagraph = category
    ? `          <w:p>\n` +
      `            <w:r>\n` +
      `              <w:rPr><w:b/></w:rPr>\n` +
      `              <w:t>[${escapedCategory}]</w:t>\n` +
      `            </w:r>\n` +
      `          </w:p>\n`
    : "";

  const commentBody = categoryParagraph + justificationParagraphs;

  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">',

    // Part 1: Package relationships
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

    // Part 2: Document relationships (links to comments.xml)
    '  <pkg:part pkg:name="/word/_rels/document.xml.rels"',
    '    pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">',
    "    <pkg:xmlData>",
    '      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '        <Relationship Id="rId1"',
    '          Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"',
    '          Target="comments.xml"/>',
    "      </Relationships>",
    "    </pkg:xmlData>",
    "  </pkg:part>",

    // Part 3: Document content (tracked change + comment anchors)
    '  <pkg:part pkg:name="/word/document.xml"',
    '    pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">',
    "    <pkg:xmlData>",
    '      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    "        <w:body>",
    "          <w:p>",
    '            <w:commentRangeStart w:id="0"/>',
    changeBody +
    '            <w:commentRangeEnd w:id="0"/>',
    "            <w:r>",
    '              <w:commentReference w:id="0"/>',
    "            </w:r>",
    "          </w:p>",
    "        </w:body>",
    "      </w:document>",
    "    </pkg:xmlData>",
    "  </pkg:part>",

    // Part 4: Comments (formatted justification)
    '  <pkg:part pkg:name="/word/comments.xml"',
    '    pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml">',
    "    <pkg:xmlData>",
    '      <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    `        <w:comment w:id="0" w:author="Stylistic" w:initials="St" w:date="${now}">`,
    commentBody,
    "        </w:comment>",
    "      </w:comments>",
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
