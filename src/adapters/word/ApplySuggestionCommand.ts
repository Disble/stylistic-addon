/* global Word, console, DOMParser, XMLSerializer */

/**
 * ApplySuggestionCommand — Command pattern for tracked-change insertion.
 *
 * Encapsulates the complete logic for applying a single `Suggestion` as a
 * native Word tracked change with an embedded justification comment.
 *
 * Each command:
 * 1. Searches the document for the original text (case-sensitive).
 * 2. Extracts run formatting (`<w:rPr>`) from the matched range.
 * 3. Disables `changeTrackingMode` (to avoid double-tracking the OOXML insertion).
 * 4. Builds the OOXML package via `OoxmlPackageBuilder`.
 * 5. Inserts the package, replacing the matched range.
 * 6. Restores `changeTrackingMode` in a `finally` block.
 *
 * Each suggestion runs in its own `Word.run` context (per-suggestion isolation)
 * to avoid stale ranges after OOXML insertions shift document positions.
 *
 * An `undo()` method can be added in a future iteration to support reverting
 * individual tracked changes.
 *
 * @module ApplySuggestionCommand
 */

import { Suggestion, CommandResult, ChangeType } from "../../domain/types";
import { OoxmlPackageBuilder } from "./ooxml/OoxmlPackageBuilder";

/**
 * Determines the type of tracked change operation for a suggestion.
 * Strategy pattern: selects insert / delete / replace based on text content.
 */
function classifyChange(suggestion: Suggestion): ChangeType {
  const hasOriginal = suggestion.originalText.length > 0;
  const hasSuggested = suggestion.suggestedText.length > 0;
  if (hasOriginal && !hasSuggested) return "delete";
  if (!hasOriginal && hasSuggested) return "insert";
  return "replace";
}

/**
 * Extracts the first `<w:rPr>` element from a flat OPC OOXML string.
 * Used to preserve the original text's formatting in the tracked change.
 */
function extractRunProperties(ooxml: string): string | null {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(ooxml, "application/xml");
    const nsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const rPr = doc.getElementsByTagNameNS(nsW, "rPr")[0];
    if (!rPr) return null;
    return new XMLSerializer().serializeToString(rPr);
  } catch {
    return null;
  }
}

/**
 * A Command that applies one `Suggestion` as a tracked change in Word.
 *
 * Implements `DocumentCommand` from `domain/types.ts`.
 */
export class ApplySuggestionCommand {
  readonly id: string;
  readonly description: string;

  constructor(private readonly suggestion: Suggestion) {
    this.id = suggestion.id;
    this.description = `Apply suggestion: "${suggestion.originalText.substring(0, 40)}" → "${suggestion.suggestedText.substring(0, 40)}"`;
  }

  /**
   * Executes the command: searches for `originalText` in the document and
   * replaces it with an OOXML tracked change package.
   *
   * Returns `{ success: false }` for recoverable application failures such as
   * missing anchor text, missing search matches, or Office insertion errors.
   */
  async execute(): Promise<CommandResult> {
    console.log(
      `🔍 [ApplySuggestionCommand] "${this.id}": "${this.suggestion.originalText.substring(0, 40)}" → "${this.suggestion.suggestedText.substring(0, 40)}"`
    );

    const changeType = classifyChange(this.suggestion);

    if (changeType === "insert") {
      console.warn(`⚠️ [ApplySuggestionCommand] "${this.id}": inserción sin texto ancla`);
      return {
        success: false,
        commandId: this.id,
        error: "Insert-only suggestions require anchor text",
      };
    }

    try {
      return await Word.run(async (context) => {
        const results = context.document.body.search(this.suggestion.originalText, {
          matchCase: true,
          matchWholeWord: false,
        });
        results.load("items");
        await context.sync();

        if (results.items.length === 0) {
          console.warn(`🔍 [ApplySuggestionCommand] "${this.id}": texto no encontrado`);
          return { success: false, commandId: this.id, error: "Texto original no encontrado" };
        }

        const range = results.items[0];
        const rangeOoxml = range.getOoxml();
        await context.sync();
        // eslint-disable-next-line office-addins/load-object-before-read
        const runProps = extractRunProperties(rangeOoxml.value);

        context.document.load("changeTrackingMode");
        await context.sync();
        const previousMode = context.document.changeTrackingMode;
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        await context.sync();

        try {
          const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");

          const ooxml = new OoxmlPackageBuilder()
            .withRunProperties(runProps)
            .withChange(
              this.suggestion.originalText,
              this.suggestion.suggestedText,
              changeType,
              "Stylistic",
              now
            )
            .withComment(this.suggestion.category, this.suggestion.justification, "Stylistic", now)
            .build();

          console.log(
            `📄 [ApplySuggestionCommand] "${this.id}": insertando OOXML (tipo: ${changeType})`
          );
          range.insertOoxml(ooxml, Word.InsertLocation.replace);
          await context.sync();
          console.log(`✅ [ApplySuggestionCommand] "${this.id}": insertado exitosamente`);

          return { success: true, commandId: this.id };
        } finally {
          context.document.changeTrackingMode = previousMode as Word.ChangeTrackingMode;
          await context.sync();
        }
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(`⚠️ [ApplySuggestionCommand] "${this.id}": ${message}`);
      return { success: false, commandId: this.id, error: message };
    }
  }
}
