import type { ParagraphSnapshot } from "./WordTextSourceAdapter.types";

/** Returns whether the Word built-in style should be treated as a heading. */
export function isHeadingStyle(styleBuiltIn?: string): boolean {
  return styleBuiltIn === "Title" || /^Heading\d+$/.test(styleBuiltIn ?? "");
}

/** Returns whether the paragraph text should preserve a leading tab for indentation. */
export function shouldPrefixIndent(paragraph: ParagraphSnapshot): boolean {
  if (isHeadingStyle(paragraph.styleBuiltIn)) {
    return false;
  }

  return (paragraph.firstLineIndent ?? 0) > 0 || (paragraph.leftIndent ?? 0) > 0;
}

/** Builds structured text with paragraph spacing and indentation preserved. */
export function buildStructuredParagraphText(paragraphs: ParagraphSnapshot[]): string {
  if (paragraphs.length === 0) {
    return "";
  }

  return paragraphs
    .map((paragraph) => {
      const text = paragraph.text ?? "";
      if (!shouldPrefixIndent(paragraph) || text.length === 0 || text.startsWith("\t")) {
        return text;
      }

      return `\t${text}`;
    })
    .join("\n\n");
}
