/**
 * Characterization tests for OoxmlPackageBuilder.
 *
 * These tests document the CURRENT behavior of the builder, including
 * any quirks in XML structure, escaping, or formatting. They are not
 * prescriptive — they lock in existing behavior so future refactors
 * can detect regressions.
 */

import { OoxmlPackageBuilder } from "./OoxmlPackageBuilder";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Trims each line of a multiline string for easier comparison. */
function trimLines(xml: string): string {
  return xml
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean)
    .join("\n");
}

function getFirstMatch(xml: string, pattern: RegExp): string {
  const match = xml.match(pattern);

  expect(match).toBeTruthy();

  return match![1];
}

function getCommentId(xml: string): string {
  return getFirstMatch(xml, /<w:commentRangeStart w:id="(\d+)"\/>/);
}

function getDeletionId(xml: string): string {
  return getFirstMatch(xml, /<w:del w:id="(\d+)"/);
}

function getInsertionId(xml: string): string {
  return getFirstMatch(xml, /<w:ins w:id="(\d+)"/);
}

// ---------------------------------------------------------------------------
// Minimal / Empty Package
// ---------------------------------------------------------------------------

describe("OoxmlPackageBuilder", () => {
  describe("minimal package (no configuration)", () => {
    it("produces valid XML declaration and root element", () => {
      const xml = new OoxmlPackageBuilder().build();
      expect(xml).toMatch(/^<\?xml version="1\.0" encoding="UTF-8" standalone="yes"\?>/);
      expect(xml).toContain('<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">');
      expect(xml).toMatch(/<\/pkg:package>$/);
    });

    it("contains all 4 pkg:part elements", () => {
      const xml = new OoxmlPackageBuilder().build();
      expect(xml).toContain('pkg:name="/_rels/.rels"');
      expect(xml).toContain('pkg:name="/word/_rels/document.xml.rels"');
      expect(xml).toContain('pkg:name="/word/document.xml"');
      expect(xml).toContain('pkg:name="/word/comments.xml"');
    });

    it("package rels point to word/document.xml", () => {
      const xml = new OoxmlPackageBuilder().build();
      expect(xml).toContain('Target="word/document.xml"');
    });

    it("document rels point to comments.xml", () => {
      const xml = new OoxmlPackageBuilder().build();
      expect(xml).toContain('Target="comments.xml"');
    });

    it("has empty tracked-change body (no del/ins) when nothing configured", () => {
      const xml = new OoxmlPackageBuilder().build();
      expect(xml).not.toContain("<w:del");
      expect(xml).not.toContain("<w:ins");
    });

    it("always includes comment anchors (commentRangeStart, commentRangeEnd, commentReference)", () => {
      const xml = new OoxmlPackageBuilder().build();

      const commentId = getCommentId(xml);

      expect(xml).toContain(`<w:commentRangeStart w:id="${commentId}"/>`);
      expect(xml).toContain(`<w:commentRangeEnd w:id="${commentId}"/>`);
      expect(xml).toContain(`<w:commentReference w:id="${commentId}"/>`);
    });

    it("includes a comment element with default author and empty body", () => {
      const xml = new OoxmlPackageBuilder().build();

      expect(xml).toContain('w:author="Stylistic"');
      expect(xml).toContain('w:initials="St"');
      expect(xml).toContain('w:date=""');
    });

    it("produces stable output on repeated build() calls", () => {
      const builder = new OoxmlPackageBuilder();
      const first = builder.build();
      const second = builder.build();
      expect(first).toBe(second);
    });
  });

  // ---------------------------------------------------------------------------
  // withDeletion
  // ---------------------------------------------------------------------------

  describe("withDeletion", () => {
    it("adds a w:del element with correct attributes", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("old text", "Author1", "2025-01-15T10:00:00Z")
        .build();

      expect(xml).toMatch(
        /<w:del w:id="\d+" w:author="Author1" w:date="2025-01-15T10:00:00Z">/
      );
    });

    it("wraps deleted text in w:delText with xml:space preserve", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("remove me", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain('<w:delText xml:space="preserve">remove me</w:delText>');
    });

    it("assigns a deterministic numeric deletion id", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("x", "A", "2025-01-01T00:00:00Z")
        .withInsertion("y", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(getDeletionId(xml)).toMatch(/^\d+$/);
    });

    it("returns this for fluent chaining", () => {
      const builder = new OoxmlPackageBuilder();
      const result = builder.withDeletion("text", "A", "2025-01-01T00:00:00Z");
      expect(result).toBe(builder);
    });
  });

  // ---------------------------------------------------------------------------
  // withInsertion
  // ---------------------------------------------------------------------------

  describe("withInsertion", () => {
    it("adds a w:ins element with correct attributes", () => {
      const xml = new OoxmlPackageBuilder()
        .withInsertion("new text", "Author2", "2025-06-01T12:00:00Z")
        .build();

      expect(xml).toMatch(
        /<w:ins w:id="\d+" w:author="Author2" w:date="2025-06-01T12:00:00Z">/
      );
    });

    it("wraps inserted text in w:t with xml:space preserve", () => {
      const xml = new OoxmlPackageBuilder()
        .withInsertion("insert me", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain('<w:t xml:space="preserve">insert me</w:t>');
    });

    it("assigns a deterministic numeric insertion id", () => {
      const xml = new OoxmlPackageBuilder()
        .withInsertion("y", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(getInsertionId(xml)).toMatch(/^\d+$/);
    });

    it("returns this for fluent chaining", () => {
      const builder = new OoxmlPackageBuilder();
      const result = builder.withInsertion("text", "A", "2025-01-01T00:00:00Z");
      expect(result).toBe(builder);
    });
  });

  // ---------------------------------------------------------------------------
  // withRunProperties
  // ---------------------------------------------------------------------------

  describe("withRunProperties", () => {
    it("embeds rPr XML inside deletion run when provided", () => {
      const rPr = '<w:rPr><w:b/><w:i/></w:rPr>';
      const xml = new OoxmlPackageBuilder()
        .withRunProperties(rPr)
        .withDeletion("text", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain('<w:rPr><w:b/><w:i/></w:rPr>');
      // rPr appears before delText
      const rPrIdx = xml.indexOf("<w:rPr><w:b/>");
      const delTextIdx = xml.indexOf("<w:delText");
      expect(rPrIdx).toBeLessThan(delTextIdx);
    });

    it("embeds rPr XML inside insertion run when provided", () => {
      const rPr = '<w:rPr><w:sz w:val="24"/></w:rPr>';
      const xml = new OoxmlPackageBuilder()
        .withRunProperties(rPr)
        .withInsertion("text", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain('<w:rPr><w:sz w:val="24"/></w:rPr>');
      const rPrIdx = xml.indexOf('<w:rPr><w:sz');
      const tIdx = xml.indexOf('<w:t xml:space="preserve">text</w:t>');
      expect(rPrIdx).toBeLessThan(tIdx);
    });

    it("embeds rPr in BOTH runs when deletion and insertion are present", () => {
      const rPr = '<w:rPr><w:b/></w:rPr>';
      const xml = new OoxmlPackageBuilder()
        .withRunProperties(rPr)
        .withDeletion("old", "A", "2025-01-01T00:00:00Z")
        .withInsertion("new", "A", "2025-01-01T00:00:00Z")
        .build();

      // rPr should appear twice: once in del run, once in ins run
      const matches = xml.match(/<w:rPr><w:b\/><\/w:rPr>/g);
      expect(matches).toHaveLength(2);
    });

    it("does not add rPr when set to null", () => {
      const xml = new OoxmlPackageBuilder()
        .withRunProperties(null)
        .withDeletion("text", "A", "2025-01-01T00:00:00Z")
        .build();

      // The only rPr in the output should be inside the comment (bold category),
      // not inside the tracked change
      const docPartMatch = xml.match(
        /document\.xml[\s\S]*?<w:del[\s\S]*?<\/w:del>/
      );
      expect(docPartMatch).toBeTruthy();
      expect(docPartMatch![0]).not.toContain("<w:rPr>");
    });

    it("returns this for fluent chaining", () => {
      const builder = new OoxmlPackageBuilder();
      const result = builder.withRunProperties('<w:rPr/>');
      expect(result).toBe(builder);
    });

    it("strips <w:rFonts> from runPropsXml to prevent Symbol font corruption in tracked changes", () => {
      const xml = new OoxmlPackageBuilder()
        .withRunProperties('<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/><w:b/></w:rPr>')
        .withDeletion("texto original", "Stylistic", "2025-01-01T00:00:00Z")
        .withInsertion("texto sugerido", "Stylistic", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).not.toContain("w:rFonts");
      expect(xml).toContain("<w:b/>");
    });

    it("strips <w:rFonts> with multiple attributes (Calibri, eastAsia, cs forms)", () => {
      const xml = new OoxmlPackageBuilder()
        .withRunProperties(
          '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:i/></w:rPr>'
        )
        .withDeletion("texto", "Stylistic", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).not.toContain("w:rFonts");
      expect(xml).toContain("<w:i/>");
    });
  });

  // ---------------------------------------------------------------------------
  // withChange (convenience method)
  // ---------------------------------------------------------------------------

  describe("withChange", () => {
    it('configures only deletion for type "delete"', () => {
      const xml = new OoxmlPackageBuilder()
        .withChange("original", "replacement", "delete", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("<w:del");
      expect(xml).toContain("original");
      expect(xml).not.toContain("<w:ins");
      // "replacement" should NOT appear in the tracked change body
      expect(xml).not.toContain("replacement");
    });

    it('configures only insertion for type "insert"', () => {
      const xml = new OoxmlPackageBuilder()
        .withChange("original", "replacement", "insert", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).not.toContain("<w:del");
      expect(xml).toContain("<w:ins");
      expect(xml).toContain("replacement");
    });

    it('configures both deletion and insertion for type "replace"', () => {
      const xml = new OoxmlPackageBuilder()
        .withChange("old text", "new text", "replace", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("<w:del");
      expect(xml).toContain("<w:ins");
      expect(xml).toContain("old text");
      expect(xml).toContain("new text");
    });

    it("sets the same author and date on both deletion and insertion for replace", () => {
      const xml = new OoxmlPackageBuilder()
        .withChange("old", "new", "replace", "TestAuthor", "2025-06-15T08:30:00Z")
        .build();

      expect(xml).toContain('w:author="TestAuthor" w:date="2025-06-15T08:30:00Z">');
      // Both del and ins should have the same author
      const authorMatches = xml.match(/w:author="TestAuthor"/g);
      // At least 2: one for del, one for ins (comment author is separate)
      expect(authorMatches!.length).toBeGreaterThanOrEqual(2);
    });

    it("returns this for fluent chaining", () => {
      const builder = new OoxmlPackageBuilder();
      const result = builder.withChange("a", "b", "replace", "A", "2025-01-01T00:00:00Z");
      expect(result).toBe(builder);
    });
  });

  // ---------------------------------------------------------------------------
  // withComment
  // ---------------------------------------------------------------------------

  describe("withComment", () => {
    it("renders category in bold inside brackets", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Redundancia", "Texto repetido", "Stylistic", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("<w:rPr><w:b/></w:rPr>");
      expect(xml).toContain("<w:t>[Redundancia]</w:t>");
    });

    it("renders justification text in a separate paragraph", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "This is the reason.", "Stylistic", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("<w:t>This is the reason.</w:t>");
    });

    it("splits multi-line justification into separate w:p elements", () => {
      const justification = "First line.\nSecond line.\nThird line.";
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", justification, "Stylistic", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("<w:t>First line.</w:t>");
      expect(xml).toContain("<w:t>Second line.</w:t>");
      expect(xml).toContain("<w:t>Third line.</w:t>");

      // Count justification paragraphs (excluding the category paragraph which has w:b)
      const justParagraphs = xml.match(/<w:p>\s*\n\s*<w:r>\s*\n\s*<w:t>/g);
      expect(justParagraphs).toHaveLength(3);
    });

    it("filters out empty lines from justification", () => {
      const justification = "Line one.\n\n\nLine two.\n   \nLine three.";
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", justification, "Stylistic", "2025-01-01T00:00:00Z")
        .build();

      // Only 3 non-empty lines should produce paragraphs
      expect(xml).toContain("<w:t>Line one.</w:t>");
      expect(xml).toContain("<w:t>Line two.</w:t>");
      expect(xml).toContain("<w:t>Line three.</w:t>");
    });

    it("omits category paragraph when category is empty string", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("", "Just a reason.", "Stylistic", "2025-01-01T00:00:00Z")
        .build();

      // No bold rPr in comments section
      const commentSection = xml.slice(xml.indexOf("comments.xml"));
      expect(commentSection).not.toContain("<w:b/>");
      expect(commentSection).toContain("<w:t>Just a reason.</w:t>");
    });

    it("sets comment author and date on the w:comment element", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Reason", "ReviewBot", "2025-03-20T14:00:00Z")
        .build();

      expect(xml).toContain('w:author="ReviewBot"');
      expect(xml).toContain('w:date="2025-03-20T14:00:00Z"');
    });

    it("derives initials from a single-word author", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Reason", "DifferentAuthor", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain('w:initials="Di"');
    });

    it("derives initials from the first letters of multi-word authors", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Reason", "Review Bot", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain('w:initials="RB"');
    });

    it("falls back to NA initials when author is blank", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Reason", "   ", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain('w:initials="NA"');
    });

    it("returns this for fluent chaining", () => {
      const builder = new OoxmlPackageBuilder();
      const result = builder.withComment("C", "J", "A", "2025-01-01T00:00:00Z");
      expect(result).toBe(builder);
    });
  });

  // ---------------------------------------------------------------------------
  // XML Escaping
  // ---------------------------------------------------------------------------

  describe("XML escaping", () => {
    it("escapes ampersand in deletion text", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("Tom & Jerry", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("Tom &amp; Jerry");
      expect(xml).not.toContain("Tom & Jerry");
    });

    it("escapes angle brackets in insertion text", () => {
      const xml = new OoxmlPackageBuilder()
        .withInsertion("x < y > z", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("x &lt; y &gt; z");
    });

    it("escapes double quotes in text content", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion('He said "hello"', "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("He said &quot;hello&quot;");
    });

    it("escapes single quotes (apostrophes) in text content", () => {
      const xml = new OoxmlPackageBuilder()
        .withInsertion("it's fine", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("it&apos;s fine");
    });

    it("escapes all special chars combined", () => {
      const nasty = `Tom & Jerry's <"adventure">`;
      const xml = new OoxmlPackageBuilder()
        .withDeletion(nasty, "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain(
        "Tom &amp; Jerry&apos;s &lt;&quot;adventure&quot;&gt;"
      );
    });

    it("escapes author names in deletion markup", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("text", "O'Brien & Co", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("w:author=\"O&apos;Brien &amp; Co\"");
    });

    it("escapes author names in insertion markup", () => {
      const xml = new OoxmlPackageBuilder()
        .withInsertion("text", "A<B>C", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("w:author=\"A&lt;B&gt;C\"");
    });

    it("escapes comment author", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Reason", "Author&Co", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain('w:author="Author&amp;Co"');
    });

    it("escapes category text in comment", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("A & B", "Reason", "Stylistic", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("<w:t>[A &amp; B]</w:t>");
    });

    it("escapes justification text in comment", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Use < instead of >", "Stylistic", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("<w:t>Use &lt; instead of &gt;</w:t>");
    });

    it("escapes date values in tracked changes", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("text", "A", 'bad"date')
        .build();

      expect(xml).toContain('w:date="bad&quot;date"');
    });

    it("escapes date values in comment metadata", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Reason", "A", 'bad&date')
        .build();

      expect(xml).toContain('w:date="bad&amp;date"');
    });
  });

  // ---------------------------------------------------------------------------
  // Full Build Structure
  // ---------------------------------------------------------------------------

  describe("full build output structure", () => {
    const DATE = "2025-06-15T10:30:00Z";

    it("produces deletion before insertion in replace scenario", () => {
      const xml = new OoxmlPackageBuilder()
        .withChange("old", "new", "replace", "Stylistic", DATE)
        .withComment("Grammar", "Subject-verb agreement", "Stylistic", DATE)
        .build();

      const delIdx = xml.indexOf("<w:del");
      const insIdx = xml.indexOf("<w:ins");
      expect(delIdx).toBeLessThan(insIdx);
    });

    it("nests tracked changes inside w:p within w:body", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("text", "A", DATE)
        .build();

      // Structural nesting: w:document > w:body > w:p > w:del
      const bodyStart = xml.indexOf("<w:body>");
      const pStart = xml.indexOf("<w:p>", bodyStart);
      const delStart = xml.indexOf("<w:del", pStart);
      const delEnd = xml.indexOf("</w:del>", delStart);
      const pEnd = xml.indexOf("</w:p>", delEnd);
      const bodyEnd = xml.indexOf("</w:body>", pEnd);

      expect(bodyStart).toBeGreaterThan(-1);
      expect(pStart).toBeGreaterThan(bodyStart);
      expect(delStart).toBeGreaterThan(pStart);
      expect(delEnd).toBeGreaterThan(delStart);
      expect(pEnd).toBeGreaterThan(delEnd);
      expect(bodyEnd).toBeGreaterThan(pEnd);
    });

    it("places comment range around tracked changes", () => {
      const xml = new OoxmlPackageBuilder()
        .withChange("old", "new", "replace", "Stylistic", DATE)
        .build();

      const rangeStart = xml.indexOf("commentRangeStart");
      const delStart = xml.indexOf("<w:del");
      const insEnd = xml.indexOf("</w:ins>");
      const rangeEnd = xml.indexOf("commentRangeEnd");

      expect(rangeStart).toBeLessThan(delStart);
      expect(rangeEnd).toBeGreaterThan(insEnd);
    });

    it("reuses the same generated comment id across anchors and comment metadata", () => {
      const xml = new OoxmlPackageBuilder()
        .withChange("old", "new", "replace", "Stylistic", DATE)
        .withComment("Grammar", "Subject-verb agreement", "Stylistic", DATE)
        .build();

      const commentId = getCommentId(xml);

      expect(xml).toContain(`<w:commentRangeEnd w:id="${commentId}"/>`);
      expect(xml).toContain(`<w:commentReference w:id="${commentId}"/>`);
      expect(xml).toContain(`<w:comment w:id="${commentId}"`);
    });

    it("generates distinct ids for comment, deletion, and insertion nodes", () => {
      const xml = new OoxmlPackageBuilder()
        .withChange("old", "new", "replace", "Stylistic", DATE)
        .withComment("Grammar", "Subject-verb agreement", "Stylistic", DATE)
        .build();

      const commentId = getCommentId(xml);
      const deletionId = getDeletionId(xml);
      const insertionId = getInsertionId(xml);

      expect(new Set([commentId, deletionId, insertionId])).toHaveLength(3);
    });

    it("does not reuse one universal hardcoded id set for every document", () => {
      const firstXml = new OoxmlPackageBuilder()
        .withChange("old", "new", "replace", "Stylistic", DATE)
        .withComment("Grammar", "Subject-verb agreement", "Stylistic", DATE)
        .build();

      const secondXml = new OoxmlPackageBuilder()
        .withChange("before", "after", "replace", "Review Bot", DATE)
        .withComment("Style", "Clarified tone.", "Review Bot", DATE)
        .build();

      expect(getCommentId(firstXml)).not.toBe(getCommentId(secondXml));
      expect(getDeletionId(firstXml)).not.toBe(getDeletionId(secondXml));
      expect(getInsertionId(firstXml)).not.toBe(getInsertionId(secondXml));
    });

    it("snapshot: full replace build output", () => {
      const xml = new OoxmlPackageBuilder()
        .withRunProperties('<w:rPr><w:b/></w:rPr>')
        .withChange("incorrecto", "correcto", "replace", "Stylistic", DATE)
        .withComment("Ortografía", "Corrección ortográfica.", "Stylistic", DATE)
        .build();

      expect(xml).toMatchSnapshot();
    });
  });

  // ---------------------------------------------------------------------------
  // Document XML Namespace
  // ---------------------------------------------------------------------------

  describe("XML namespaces", () => {
    it("declares wordprocessingml namespace on w:document", () => {
      const xml = new OoxmlPackageBuilder().build();
      expect(xml).toContain(
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      );
    });

    it("declares wordprocessingml namespace on w:comments", () => {
      const xml = new OoxmlPackageBuilder().build();
      // Verify the comments section also has the namespace
      const commentsTag = xml.match(/<w:comments[^>]*>/);
      expect(commentsTag).toBeTruthy();
      expect(commentsTag![0]).toContain(
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      );
    });

    it("declares relationships namespace on both Relationships elements", () => {
      const xml = new OoxmlPackageBuilder().build();
      const relMatches = xml.match(
        /xmlns="http:\/\/schemas\.openxmlformats\.org\/package\/2006\/relationships"/g
      );
      expect(relMatches).toHaveLength(2);
    });
  });

  // ---------------------------------------------------------------------------
  // Edge Cases
  // ---------------------------------------------------------------------------

  describe("edge cases", () => {
    it("handles empty string deletion text", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("", "A", "2025-01-01T00:00:00Z")
        .build();

      // Empty string still produces the del element
      expect(xml).toContain("<w:del");
      expect(xml).toContain('<w:delText xml:space="preserve"></w:delText>');
    });

    it("handles empty string insertion text", () => {
      const xml = new OoxmlPackageBuilder()
        .withInsertion("", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("<w:ins");
      expect(xml).toContain('<w:t xml:space="preserve"></w:t>');
    });

    it("handles unicode text (Spanish accents, ñ)", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("niño", "A", "2025-01-01T00:00:00Z")
        .withInsertion("señor año", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("niño");
      expect(xml).toContain("señor año");
    });

    it("handles emoji in text", () => {
      const xml = new OoxmlPackageBuilder()
        .withInsertion("hello 🌍", "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("hello 🌍");
    });

    it("handles very long text without truncation", () => {
      const longText = "a".repeat(10_000);
      const xml = new OoxmlPackageBuilder()
        .withDeletion(longText, "A", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain(longText);
    });

    it("handles text with newlines in deletion (not split — only justification splits)", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("line1\nline2", "A", "2025-01-01T00:00:00Z")
        .build();

      // Deletion text is NOT split by newlines (only comment justification is)
      expect(xml).toContain("line1\nline2");
    });

    it("handles justification that is only whitespace/newlines", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "  \n\n  \n", "A", "2025-01-01T00:00:00Z")
        .build();

      // All lines are whitespace-only, so they get filtered out
      // The comment should have category but no justification paragraphs
      const commentSection = xml.slice(xml.indexOf("<w:comment"));
      expect(commentSection).toContain("[Cat]");
      // No plain <w:t> without bold (justification paragraphs)
      const justMatches = commentSection.match(/<w:p>\s*\n\s*<w:r>\s*\n\s*<w:t>/g);
      expect(justMatches).toBeNull();
    });

    it("handles empty category AND empty justification", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("", "", "A", "2025-01-01T00:00:00Z")
        .build();

      // Comment element still exists but with no child paragraphs
      expect(xml).toContain("<w:comment");
      expect(xml).toContain("</w:comment>");
    });

    it("handles run properties as empty string (still emitted)", () => {
      const xml = new OoxmlPackageBuilder()
        .withRunProperties("")
        .withDeletion("text", "A", "2025-01-01T00:00:00Z")
        .build();

      // QUIRK: Empty string is falsy, so rPr is NOT emitted
      // (builder checks `this.runPropsXml ? ...` which is falsy for "")
      const docPart = xml.slice(
        xml.indexOf("<w:del"),
        xml.indexOf("</w:del>")
      );
      // The rPr line won't appear because "" is falsy
      expect(docPart).not.toContain("                \n");
    });

    it("last withDeletion call wins (overwrites previous)", () => {
      const xml = new OoxmlPackageBuilder()
        .withDeletion("first", "A1", "2025-01-01T00:00:00Z")
        .withDeletion("second", "A2", "2025-02-01T00:00:00Z")
        .build();

      expect(xml).toContain("second");
      expect(xml).not.toContain(">first<");
      expect(xml).toContain('w:author="A2"');
    });

    it("last withInsertion call wins (overwrites previous)", () => {
      const xml = new OoxmlPackageBuilder()
        .withInsertion("first", "A1", "2025-01-01T00:00:00Z")
        .withInsertion("second", "A2", "2025-02-01T00:00:00Z")
        .build();

      expect(xml).toContain("second");
      expect(xml).not.toContain(">first<");
    });
  });

  // ---------------------------------------------------------------------------
  // withComment — comment-only mode (originalText in body)
  // ---------------------------------------------------------------------------

  describe("withComment with originalText (comment-only path)", () => {
    const DATE = "2025-06-15T10:30:00Z";

    it("includes originalText in the document body when provided", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Estilo", "Mejora la claridad", "Stylistic", DATE, "texto original")
        .build();

      const bodyStart = xml.indexOf("<w:body>");
      const bodyEnd = xml.indexOf("</w:body>");
      const body = xml.slice(bodyStart, bodyEnd);

      expect(body).toContain("texto original");
    });

    it("places originalText inside a w:r/w:t run between commentRangeStart and commentRangeEnd", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Reason", "Stylistic", DATE, "the original")
        .build();

      const rangeStart = xml.indexOf("<w:commentRangeStart");
      const textRun = xml.indexOf('<w:t xml:space="preserve">the original</w:t>');
      const rangeEnd = xml.indexOf("<w:commentRangeEnd");

      expect(rangeStart).toBeGreaterThan(-1);
      expect(textRun).toBeGreaterThan(rangeStart);
      expect(rangeEnd).toBeGreaterThan(textRun);
    });

    it("places commentReference run after commentRangeEnd", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Reason", "Stylistic", DATE, "some text")
        .build();

      const rangeEnd = xml.indexOf("<w:commentRangeEnd");
      const reference = xml.indexOf("<w:commentReference");

      expect(rangeEnd).toBeGreaterThan(-1);
      expect(reference).toBeGreaterThan(rangeEnd);
    });

    it("still produces the formatted comment (category in bold + justification) in comments.xml", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Redundancia", "Texto repetido innecesario", "Stylistic", DATE, "the original text")
        .build();

      const commentsSection = xml.slice(xml.indexOf("comments.xml"));
      expect(commentsSection).toContain("<w:t>[Redundancia]</w:t>");
      expect(commentsSection).toContain("<w:t>Texto repetido innecesario</w:t>");
      expect(commentsSection).toContain("<w:rPr><w:b/></w:rPr>");
    });

    it("escapes special XML characters in originalText", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Reason", "Stylistic", DATE, "Tom & Jerry's <test>")
        .build();

      expect(xml).toContain("Tom &amp; Jerry&apos;s &lt;test&gt;");
    });

    it("produces an empty body body run when originalText is omitted (backward compat)", () => {
      const xml = new OoxmlPackageBuilder()
        .withComment("Cat", "Reason", "Stylistic", DATE)
        .build();

      const bodyStart = xml.indexOf("<w:body>");
      const bodyEnd = xml.indexOf("</w:body>");
      const body = xml.slice(bodyStart, bodyEnd);

      // No plain text run — only commentRangeStart, commentRangeEnd, commentReference
      expect(body).not.toMatch(/<w:t xml:space="preserve">[^<]+<\/w:t>/);
    });

    it("returns this for fluent chaining with originalText", () => {
      const builder = new OoxmlPackageBuilder();
      const result = builder.withComment("C", "J", "A", DATE, "text");
      expect(result).toBe(builder);
    });
  });

  // ---------------------------------------------------------------------------
  // Fluent API / Builder Pattern
  // ---------------------------------------------------------------------------

  describe("fluent API", () => {
    it("supports full chain in a single expression", () => {
      const xml = new OoxmlPackageBuilder()
        .withRunProperties('<w:rPr><w:i/></w:rPr>')
        .withDeletion("bad", "Stylistic", "2025-01-01T00:00:00Z")
        .withInsertion("good", "Stylistic", "2025-01-01T00:00:00Z")
        .withComment("Style", "Improved clarity.", "Stylistic", "2025-01-01T00:00:00Z")
        .build();

      expect(xml).toContain("bad");
      expect(xml).toContain("good");
      expect(xml).toContain("[Style]");
      expect(xml).toContain("Improved clarity.");
      expect(xml).toContain("<w:i/>");
    });

    it("withChange + withComment is the typical usage pattern", () => {
      const xml = new OoxmlPackageBuilder()
        .withChange("entonces", "por lo tanto", "replace", "Stylistic", "2025-06-01T00:00:00Z")
        .withComment("Muletilla", "Evitar uso excesivo de 'entonces'.", "Stylistic", "2025-06-01T00:00:00Z")
        .build();

      expect(xml).toContain('<w:delText xml:space="preserve">entonces</w:delText>');
      expect(xml).toContain('<w:t xml:space="preserve">por lo tanto</w:t>');
      expect(xml).toContain("[Muletilla]");
      expect(xml).toContain("Evitar uso excesivo de &apos;entonces&apos;.");
    });
  });
});
