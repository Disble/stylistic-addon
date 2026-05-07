import type { TextSource } from "../../domain/TextSource.types";
import { PARAGRAPH_LOAD_FIELDS } from "./WordTextSourceAdapter.constants";
import { buildStructuredParagraphText } from "./WordTextSourceAdapter.helpers";
import type { ParagraphSnapshot } from "./WordTextSourceAdapter.types";

/** Reads and normalizes the Word text source for analysis. */
export class WordTextSourceAdapter {
  /** Resolves selection text first, then falls back to the full body. */
  async getTextToAnalyze(): Promise<TextSource> {
    console.log("📖 [WordAdapter] Resolviendo texto a analizar...");
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      // eslint-disable-next-line office-addins/no-navigational-load -- Word paragraph collections must load `paragraphs/items` before projecting paragraph members.
      selection.load("text,paragraphs/items");
      // eslint-disable-next-line office-addins/no-empty-load -- The Office lint parser does not resolve variable-backed load strings; `PARAGRAPH_LOAD_FIELDS` is explicit and non-empty.
      selection.paragraphs.load(PARAGRAPH_LOAD_FIELDS);
      await context.sync();

      const selectionRawText = selection.text;
      const hasSelectedText = selectionRawText.trim().length > 0;
      const selectionText = hasSelectedText
        ? buildStructuredParagraphText(selection.paragraphs.items as ParagraphSnapshot[]) ||
          selectionRawText
        : "";

      if (selectionText && selectionText.trim().length > 0) {
        console.log(`📖 [WordAdapter] Selección activa — ${selectionText.length} chars`);
        return { text: selectionText, isSelection: true };
      }

      const body = context.document.body;
      body.load("paragraphs/items");
      body.paragraphs.load(PARAGRAPH_LOAD_FIELDS);
      await context.sync();

      const bodyText = buildStructuredParagraphText(body.paragraphs.items as ParagraphSnapshot[]);
      if (bodyText.length > 0) {
        console.log(`📖 [WordAdapter] Documento completo — ${bodyText.length} chars`);
        return { text: bodyText, isSelection: false };
      }

      body.load("text");
      await context.sync();
      console.log(`📖 [WordAdapter] Documento completo — ${body.text.length} chars`);
      return { text: body.text, isSelection: false };
    });
  }
}
