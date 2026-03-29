/**
 * OOXML Package Builder — Builder pattern for flat OPC OOXML packages.
 *
 * Replaces the monolithic `buildTrackedChangeOoxml()` function with a
 * fluent builder API. Each method configures one aspect of the package.
 * `build()` assembles and returns the final XML string.
 *
 * The produced package is a flat OPC document suitable for `Range.insertOoxml()`
 * in the Office.js Word API. It contains 4 parts:
 * 1. `/_rels/.rels` — Package relationships.
 * 2. `/word/_rels/document.xml.rels` — Document relationships (links comments).
 * 3. `/word/document.xml` — Tracked change markup (`<w:del>`, `<w:ins>`).
 * 4. `/word/comments.xml` — Formatted justification comment.
 *
 * Usage:
 * ```typescript
 * const ooxml = new OoxmlPackageBuilder()
 *   .withRunProperties(runPropsXml)
 *   .withDeletion(originalText, "Stylistic", isoDate)
 *   .withInsertion(replacementText, "Stylistic", isoDate)
 *   .withComment(category, justification, "Stylistic", isoDate)
 *   .build();
 * ```
 *
 * @module OoxmlPackageBuilder
 */

import { ChangeType } from "../../../domain/types";

/** Escapes special XML characters to prevent injection in OOXML content. */
function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function hashString(value: string): number {
  let hash = 0;

  for (let index = 0; index < value.length; index += 1) {
    hash = (hash * 31 + value.charCodeAt(index)) >>> 0;
  }

  return hash;
}

function deriveInitials(author: string): string {
  const parts = author
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  if (parts.length === 0) {
    return "NA";
  }

  if (parts.length === 1) {
    const [part] = parts;
    const first = part[0]?.toUpperCase() ?? "N";
    const second = part[1]?.toLowerCase() ?? "A";

    return `${first}${second}`;
  }

  return parts
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase() ?? "")
    .join("") || "NA";
}

export class OoxmlPackageBuilder {
  private runPropsXml: string | null = null;
  private deletionText: string | null = null;
  private insertionText: string | null = null;
  private deletionAuthor = "Stylistic";
  private deletionDate = "";
  private insertionAuthor = "Stylistic";
  private insertionDate = "";
  private commentCategory = "";
  private commentJustification = "";
  private commentAuthor = "Stylistic";
  private commentDate = "";
  private commentOriginalText = "";

  private createBuildIds(): { commentId: string; deletionId: string; insertionId: string } {
    const seed = [
      this.runPropsXml ?? "",
      this.deletionText ?? "",
      this.deletionAuthor,
      this.deletionDate,
      this.insertionText ?? "",
      this.insertionAuthor,
      this.insertionDate,
      this.commentCategory,
      this.commentJustification,
      this.commentAuthor,
      this.commentDate,
    ].join("\u001f");

    const baseId = (hashString(seed) % 900000000) + 100000000;

    return {
      commentId: String(baseId),
      deletionId: String(baseId + 1),
      insertionId: String(baseId + 2),
    };
  }

  /**
   * Preserves original run formatting by embedding `<w:rPr>` XML inside
   * the tracked change markup.
   *
   * @param rPrXml - Serialized `<w:rPr>` element from the original range.
   */
  withRunProperties(rPrXml: string | null): this {
    if (rPrXml === null) {
      this.runPropsXml = null;
      return this;
    }
    // Strip <w:rFonts> to prevent Symbol/decorative font corruption:
    // the first run of a matched range may use a special font for em-dashes or
    // other typographic characters. Embedding that font in the tracked change
    // causes ALL text in the del/ins runs to render in that font (e.g., Symbol
    // maps Latin characters to Greek symbols). Font family is inherited from the
    // document context; only character-level styling (bold, italic, size…) should
    // be preserved.
    this.runPropsXml = rPrXml.replace(/<w:rFonts\b[^>]*\/?>/g, "");
    return this;
  }

  /**
   * Adds a `<w:del>` deletion node to the tracked change body.
   *
   * @param text   - The original text being deleted.
   * @param author - Author shown in the Review pane (e.g., "Stylistic").
   * @param date   - ISO 8601 timestamp for the tracked change.
   */
  withDeletion(text: string, author: string, date: string): this {
    this.deletionText = text;
    this.deletionAuthor = author;
    this.deletionDate = date;
    return this;
  }

  /**
   * Adds a `<w:ins>` insertion node to the tracked change body.
   *
   * @param text   - The replacement text being inserted.
   * @param author - Author shown in the Review pane (e.g., "Stylistic").
   * @param date   - ISO 8601 timestamp for the tracked change.
   */
  withInsertion(text: string, author: string, date: string): this {
    this.insertionText = text;
    this.insertionAuthor = author;
    this.insertionDate = date;
    return this;
  }

  /**
   * Convenience method: configures both deletion and insertion for a
   * replace-type tracked change.
   *
   * @param original    - The original text to replace.
   * @param replacement - The new text.
   * @param type        - The change type (insert/delete/replace).
   * @param author      - Author for both tracked changes.
   * @param date        - ISO 8601 timestamp.
   */
  withChange(
    original: string,
    replacement: string,
    type: ChangeType,
    author: string,
    date: string
  ): this {
    if (type === "delete" || type === "replace") {
      this.withDeletion(original, author, date);
    }
    if (type === "insert" || type === "replace") {
      this.withInsertion(replacement, author, date);
    }
    return this;
  }

  /**
   * Adds the comment (margin balloon) with bold category header and
   * justification body.
   *
   * @param category      - Editorial category label shown in bold (e.g., "Redundancia").
   * @param justification - Explanation text shown below the category.
   * @param author        - Comment author (e.g., "Stylistic").
   * @param date          - ISO 8601 timestamp for the comment.
   * @param originalText  - The original text to preserve in the document body (comment-only path).
   *                        When provided, the body includes the text wrapped in commentRange anchors
   *                        so the original text is NOT erased by `insertOoxml` replace.
   */
  withComment(category: string, justification: string, author: string, date: string, originalText = ""): this {
    this.commentCategory = category;
    this.commentJustification = justification;
    this.commentAuthor = author;
    this.commentDate = date;
    this.commentOriginalText = originalText;
    return this;
  }

  /**
   * Assembles and returns the flat OPC OOXML package as a string.
   * Safe to call multiple times (builder state is not consumed).
   */
  build(): string {
    const rPr = this.runPropsXml ? `                ${this.runPropsXml}\n` : "";
    const ids = this.createBuildIds();
    const commentInitials = deriveInitials(this.commentAuthor);

    // Build tracked change body
    let changeBody = "";

    if (this.deletionText !== null) {
      changeBody +=
        `            <w:del w:id="${ids.deletionId}" w:author="${escapeXml(this.deletionAuthor)}" w:date="${escapeXml(this.deletionDate)}">\n` +
        `              <w:r>\n` +
        `${rPr}                <w:delText xml:space="preserve">${escapeXml(this.deletionText)}</w:delText>\n` +
        `              </w:r>\n` +
        `            </w:del>\n`;
    }

    if (this.insertionText !== null) {
      changeBody +=
        `            <w:ins w:id="${ids.insertionId}" w:author="${escapeXml(this.insertionAuthor)}" w:date="${escapeXml(this.insertionDate)}">\n` +
        `              <w:r>\n` +
        `${rPr}                <w:t xml:space="preserve">${escapeXml(this.insertionText)}</w:t>\n` +
        `              </w:r>\n` +
        `            </w:ins>\n`;
    }

    // For comment-only mode: when no tracked change is present but originalText is provided,
    // emit a plain text run so insertOoxml(replace) does NOT erase the matched text.
    if (changeBody === "" && this.commentOriginalText !== "") {
      changeBody =
        `            <w:r><w:t xml:space="preserve">${escapeXml(this.commentOriginalText)}</w:t></w:r>\n`;
    }

    // Build comment body: bold category + each justification line as <w:p>
    const categoryParagraph = this.commentCategory
      ? `          <w:p>\n` +
        `            <w:r>\n` +
        `              <w:rPr><w:b/></w:rPr>\n` +
        `              <w:t>[${escapeXml(this.commentCategory)}]</w:t>\n` +
        `            </w:r>\n` +
        `          </w:p>\n`
      : "";

    const justificationParagraphs = this.commentJustification
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
      `            <w:commentRangeStart w:id="${ids.commentId}"/>`,
      changeBody + `            <w:commentRangeEnd w:id="${ids.commentId}"/>`,
      "            <w:r>",
      `              <w:commentReference w:id="${ids.commentId}"/>`,
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
      `        <w:comment w:id="${ids.commentId}" w:author="${escapeXml(this.commentAuthor)}" w:initials="${escapeXml(commentInitials)}" w:date="${escapeXml(this.commentDate)}">`,
      commentBody,
      "        </w:comment>",
      "      </w:comments>",
      "    </pkg:xmlData>",
      "  </pkg:part>",

      "</pkg:package>",
    ].join("\n");
  }
}
