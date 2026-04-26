/* global Word, console */

import type { TextSource } from "../../domain/TextSource.types";
import type { ParagraphSnapshot } from "./WordTextSourceAdapter.types";

const PARAGRAPH_LOAD_FIELDS =
  "items/text,items/styleBuiltIn,items/firstLineIndent,items/leftIndent";

function isHeadingStyle(styleBuiltIn?: string): boolean {
  return styleBuiltIn === "Title" || /^Heading\d+$/.test(styleBuiltIn ?? "");
}

function shouldPrefixIndent(paragraph: ParagraphSnapshot): boolean {
  if (isHeadingStyle(paragraph.styleBuiltIn)) {
    return false;
  }

  return (
    (paragraph.firstLineIndent ?? 0) > 0 || (paragraph.leftIndent ?? 0) > 0
  );
}

function buildStructuredParagraphText(paragraphs: ParagraphSnapshot[]): string {
  if (paragraphs.length === 0) {
    return "";
  }

  return paragraphs
    .map((paragraph) => {
      const text = paragraph.text ?? "";
      if (
        !shouldPrefixIndent(paragraph) ||
        text.length === 0 ||
        text.startsWith("\t")
      ) {
        return text;
      }

      return `\t${text}`;
    })
    .join("\n\n");
}

/** Reads and normalizes the Word text source for analysis. */
export class WordTextSourceAdapter {
  /** Resolves selection text first, then falls back to the full body. */
  async getTextToAnalyze(): Promise<TextSource> {
    console.log("📖 [WordAdapter] Resolviendo texto a analizar...");
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      selection.paragraphs.load(PARAGRAPH_LOAD_FIELDS);
      await context.sync();

      const hasSelectedText = selection.text.trim().length > 0;
      const selectionText = hasSelectedText
        ? buildStructuredParagraphText(
            selection.paragraphs.items as ParagraphSnapshot[],
          ) || selection.text
        : "";

      if (selectionText && selectionText.trim().length > 0) {
        console.log(
          `📖 [WordAdapter] Selección activa — ${selectionText.length} chars`,
        );
        return { text: selectionText, isSelection: true };
      }

      const body = context.document.body;
      body.paragraphs.load(PARAGRAPH_LOAD_FIELDS);
      await context.sync();

      const bodyText = buildStructuredParagraphText(
        body.paragraphs.items as ParagraphSnapshot[],
      );
      if (bodyText.length > 0) {
        console.log(
          `📖 [WordAdapter] Documento completo — ${bodyText.length} chars`,
        );
        return { text: bodyText, isSelection: false };
      }

      body.load("text");
      await context.sync();
      console.log(
        `📖 [WordAdapter] Documento completo — ${body.text.length} chars`,
      );
      return { text: body.text, isSelection: false };
    });
  }
}
