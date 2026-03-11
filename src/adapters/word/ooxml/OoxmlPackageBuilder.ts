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

  /**
   * Preserves original run formatting by embedding `<w:rPr>` XML inside
   * the tracked change markup.
   *
   * @param rPrXml - Serialized `<w:rPr>` element from the original range.
   */
  withRunProperties(rPrXml: string | null): this {
    this.runPropsXml = rPrXml;
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
   */
  withComment(category: string, justification: string, author: string, date: string): this {
    this.commentCategory = category;
    this.commentJustification = justification;
    this.commentAuthor = author;
    this.commentDate = date;
    return this;
  }

  /**
   * Assembles and returns the flat OPC OOXML package as a string.
   * Safe to call multiple times (builder state is not consumed).
   */
  build(): string {
    const rPr = this.runPropsXml ? `                ${this.runPropsXml}\n` : "";

    // Build tracked change body
    let changeBody = "";

    if (this.deletionText !== null) {
      changeBody +=
        `            <w:del w:id="1" w:author="${escapeXml(this.deletionAuthor)}" w:date="${this.deletionDate}">\n` +
        `              <w:r>\n` +
        `${rPr}                <w:delText xml:space="preserve">${escapeXml(this.deletionText)}</w:delText>\n` +
        `              </w:r>\n` +
        `            </w:del>\n`;
    }

    if (this.insertionText !== null) {
      changeBody +=
        `            <w:ins w:id="2" w:author="${escapeXml(this.insertionAuthor)}" w:date="${this.insertionDate}">\n` +
        `              <w:r>\n` +
        `${rPr}                <w:t xml:space="preserve">${escapeXml(this.insertionText)}</w:t>\n` +
        `              </w:r>\n` +
        `            </w:ins>\n`;
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
      '            <w:commentRangeStart w:id="0"/>',
      changeBody + '            <w:commentRangeEnd w:id="0"/>',
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
      `        <w:comment w:id="0" w:author="${escapeXml(this.commentAuthor)}" w:initials="St" w:date="${this.commentDate}">`,
      commentBody,
      "        </w:comment>",
      "      </w:comments>",
      "    </pkg:xmlData>",
      "  </pkg:part>",

      "</pkg:package>",
    ].join("\n");
  }
}
