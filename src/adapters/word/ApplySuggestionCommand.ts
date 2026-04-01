/* global Word, console */

/**
 * ApplySuggestionCommand — Command pattern for tracked-change insertion.
 *
 * Encapsulates the complete logic for applying a single `Suggestion` as a
 * native Word tracked change with an embedded justification comment.
 *
 * Each command:
 * 1. Searches the document for the original text (case-sensitive).
 * 2. Sets changeTrackingMode to trackAll.
 * 3. Calls range.insertText(suggestedText, replace) — Word records it as TC.
 * 4. Inserts a comment on the inserted range with category + justification.
 * 5. Wraps the inserted range in a ContentControl tagged `stylistic:{type}:{id}`.
 * 6. Restores `changeTrackingMode` in a `finally` block.
 *
 * Each suggestion runs in its own `Word.run` context (per-suggestion isolation)
 * to avoid stale ranges after insertions shift document positions.
 *
 * @module ApplySuggestionCommand
 */

import type { ChangeType, CommandResult, Suggestion } from "../../domain/types";

type IndexedText = {
  text: string;
  indices: number[];
};

/**
 * Determines the type of tracked change operation for a suggestion.
 * Strategy pattern: selects insert / delete / replace based on text content.
 *
 * Must only be called for `"track-change"` suggestions where `suggestedText`
 * is defined. Comment-only suggestions are handled by a separate branch in
 * `execute()` and never reach this function.
 */
function classifyChange(suggestion: Suggestion): ChangeType {
  const hasOriginal = suggestion.originalText.length > 0;
  const hasSuggested = (suggestion.suggestedText?.length ?? 0) > 0;
  if (hasOriginal && !hasSuggested) return "delete";
  if (!hasOriginal && hasSuggested) return "insert";
  return "replace";
}

function removeWhitespaceWithIndices(text: string): IndexedText {
  const indices: number[] = [];
  let normalized = "";

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (/\s/.test(char)) {
      continue;
    }

    normalized += char;
    indices.push(index);
  }

  return { text: normalized, indices };
}

function findWhitespaceInsensitiveSlice(
  searchText: string,
  documentText: string,
): string | null {
  const normalizedSearch = removeWhitespaceWithIndices(searchText).text;
  if (normalizedSearch.length === 0) {
    return null;
  }

  const normalizedDocument = removeWhitespaceWithIndices(documentText);
  const matchIndex = normalizedDocument.text.indexOf(normalizedSearch);
  if (matchIndex === -1) {
    return null;
  }

  const start = normalizedDocument.indices[matchIndex];
  const end =
    normalizedDocument.indices[matchIndex + normalizedSearch.length - 1] + 1;
  return documentText.slice(start, end);
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
    this.description = `Apply suggestion [${suggestion.type}]: "${suggestion.originalText.substring(0, 40)}" → "${(suggestion.suggestedText ?? "").substring(0, 40)}"`;
  }

  /**
   * Executes the command: searches for `originalText` in the document and
   * replaces it with a native Word tracked change.
   *
   * Returns `{ success: false }` for recoverable application failures such as
   * missing anchor text, missing search matches, or Office insertion errors.
   */
  async execute(): Promise<CommandResult> {
    console.log(
      `🔍 [ApplySuggestionCommand] "${this.id}": "${this.suggestion.originalText.substring(0, 40)}" → "${(this.suggestion.suggestedText ?? "").substring(0, 40)}"`,
    );

    if (this.suggestion.type === "comment-only") {
      return this.executeCommentOnly();
    }

    const changeType = classifyChange(this.suggestion);

    if (changeType === "insert") {
      console.warn(
        `⚠️ [ApplySuggestionCommand] "${this.id}": inserción sin texto ancla`,
      );
      return {
        success: false,
        commandId: this.id,
        error: "Insert-only suggestions require anchor text",
      };
    }

    try {
      return await Word.run(async (context) => {
        const searchOptions = {
          matchCase: true,
          matchWholeWord: false,
        };

        let searchText = this.suggestion.originalText;

        // Attempt 1: exact match (skip if text exceeds 256-char Word API limit)
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        let results!: Word.RangeCollection;
        if (searchText.length <= 256) {
          results = context.document.body.search(searchText, searchOptions);
          results.load("items");
          await context.sync();
        }

        // Attempt 1.5: ignorePunct + ignoreSpace — handles em-dashes, ¡, ¿, …
        // Triggered when text exceeds the API limit OR exact match found nothing.
        if (searchText.length > 256 || results.items.length === 0) {
          results = context.document.body.search(searchText, {
            matchCase: true,
            matchWholeWord: false,
            ignorePunct: true,
            ignoreSpace: true,
          });
          results.load("items");
          await context.sync();

          if (results.items.length > 0) {
            console.log(
              `🔍 [ApplySuggestionCommand] "${this.id}": encontrado con attempt 1.5 (ignorePunct+ignoreSpace)`,
            );
          }
        }

        if (results.items.length === 0) {
          context.document.body.load("text");
          await context.sync();

          const fallbackSearchText = findWhitespaceInsensitiveSlice(
            this.suggestion.originalText,
            context.document.body.text,
          );

          if (!fallbackSearchText) {
            console.warn(
              `🔍 [ApplySuggestionCommand] "${this.id}": texto no encontrado`,
            );
            return {
              success: false,
              commandId: this.id,
              error: "Texto original no encontrado",
            };
          }

          searchText = fallbackSearchText;
          results = context.document.body.search(searchText, searchOptions);
          results.load("items");
          await context.sync();

          if (results.items.length === 0) {
            console.warn(
              `🔍 [ApplySuggestionCommand] "${this.id}": texto no encontrado`,
            );
            return {
              success: false,
              commandId: this.id,
              error: "Texto original no encontrado",
            };
          }
        }

        let range = results.items[0];

        const parentCC = range.parentContentControlOrNullObject;
        parentCC.load("tag");
        await context.sync();

        const existingTag = parentCC.isNullObject ? "" : parentCC.tag;
        const isAlreadyCovered =
          existingTag.startsWith("stylistic:") ||
          /^chunk\d+-\d+$/.test(existingTag);

        if (isAlreadyCovered) {
          console.log(
            `♻️ [ApplySuggestionCommand] "${this.id}": CC existente detectado — eliminando wrapper y reinsertando`,
          );
          // keepContent: true = remove CC wrapper only, text stays in document
          // keepContent: false = remove CC wrapper AND DELETE its content ← DANGER
          parentCC.delete(true);
          await context.sync();

          results = context.document.body.search(searchText, searchOptions);
          results.load("items");
          await context.sync();

          if (results.items.length === 0) {
            console.warn(
              `🔍 [ApplySuggestionCommand] "${this.id}": texto no encontrado tras eliminar CC existente`,
            );
            return {
              success: false,
              commandId: this.id,
              error: "Texto no encontrado tras eliminar CC existente",
            };
          }

          range = results.items[0];
        }

        const previousMode = context.document.changeTrackingMode;
        context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await context.sync();

        try {
          console.log(
            `📄 [ApplySuggestionCommand] "${this.id}": insertando TC nativo (tipo: ${changeType})`,
          );

          const insertedRange = range.insertText(
            this.suggestion.suggestedText ?? "",
            Word.InsertLocation.replace,
          );

          insertedRange.insertComment(
            `[${this.suggestion.category}]\n${this.suggestion.justification}`,
          );

          const cc = insertedRange.insertContentControl();
          cc.tag = `stylistic:${this.suggestion.type}:${this.suggestion.id}`;
          cc.appearance = "Hidden";
          cc.cannotDelete = false;
          await context.sync();

          console.log(
            `✅ [ApplySuggestionCommand] "${this.id}": insertado exitosamente`,
          );

          return { success: true, commandId: this.id };
        } finally {
          context.document.changeTrackingMode =
            previousMode as Word.ChangeTrackingMode;
          await context.sync();
        }
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(`⚠️ [ApplySuggestionCommand] "${this.id}": ${message}`);
      return { success: false, commandId: this.id, error: message };
    }
  }

  /**
   * Executes the comment-only path: locates `originalText` in the document and
   * inserts a Word comment at that range with NO tracked change markup.
   *
   * changeTrackingMode is NOT read or modified in this path.
   *
   * The search + fallback logic mirrors the track-change path so both paths
   * share the same whitespace-insensitive anchor resolution behaviour.
   */
  private async executeCommentOnly(): Promise<CommandResult> {
    try {
      return await Word.run(async (context) => {
        const searchOptions = {
          matchCase: true,
          matchWholeWord: false,
        };

        let searchText = this.suggestion.originalText;

        // Attempt 1: exact match (skip if text exceeds 256-char Word API limit)
        // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
        let results!: Word.RangeCollection;
        if (searchText.length <= 256) {
          results = context.document.body.search(searchText, searchOptions);
          results.load("items");
          await context.sync();
        }

        // Attempt 1.5: ignorePunct + ignoreSpace — handles em-dashes, ¡, ¿, …
        // Triggered when text exceeds the API limit OR exact match found nothing.
        if (searchText.length > 256 || results.items.length === 0) {
          results = context.document.body.search(searchText, {
            matchCase: true,
            matchWholeWord: false,
            ignorePunct: true,
            ignoreSpace: true,
          });
          results.load("items");
          await context.sync();

          if (results.items.length > 0) {
            console.log(
              `🔍 [ApplySuggestionCommand] "${this.id}": encontrado con attempt 1.5 (ignorePunct+ignoreSpace) (comment-only)`,
            );
          }
        }

        if (results.items.length === 0) {
          context.document.body.load("text");
          await context.sync();

          const fallbackSearchText = findWhitespaceInsensitiveSlice(
            this.suggestion.originalText,
            context.document.body.text,
          );

          if (!fallbackSearchText) {
            console.warn(
              `🔍 [ApplySuggestionCommand] "${this.id}": texto no encontrado (comment-only)`,
            );
            return {
              success: false,
              commandId: this.id,
              error: "Texto original no encontrado",
            };
          }

          searchText = fallbackSearchText;
          results = context.document.body.search(searchText, searchOptions);
          results.load("items");
          await context.sync();

          if (results.items.length === 0) {
            console.warn(
              `🔍 [ApplySuggestionCommand] "${this.id}": texto no encontrado (comment-only)`,
            );
            return {
              success: false,
              commandId: this.id,
              error: "Texto original no encontrado",
            };
          }
        }

        let range = results.items[0];

        const parentCC = range.parentContentControlOrNullObject;
        parentCC.load("tag");
        await context.sync();

        const existingTag = parentCC.isNullObject ? "" : parentCC.tag;
        const isAlreadyCovered =
          existingTag.startsWith("stylistic:") ||
          /^chunk\d+-\d+$/.test(existingTag);

        if (isAlreadyCovered) {
          console.log(
            `♻️ [ApplySuggestionCommand] "${this.id}": CC existente detectado (comment-only) — eliminando wrapper y reinsertando`,
          );
          parentCC.delete(true);
          await context.sync();

          results = context.document.body.search(searchText, searchOptions);
          results.load("items");
          await context.sync();

          if (results.items.length === 0) {
            console.warn(
              `🔍 [ApplySuggestionCommand] "${this.id}": texto no encontrado tras eliminar CC existente (comment-only)`,
            );
            return {
              success: false,
              commandId: this.id,
              error: "Texto no encontrado tras eliminar CC existente",
            };
          }

          range = results.items[0];
        }

        console.log(
          `📄 [ApplySuggestionCommand] "${this.id}": insertando comment-only nativo`,
        );

        range.insertComment(
          `[${this.suggestion.category}]\n${this.suggestion.justification}`,
        );

        const cc = range.insertContentControl();
        cc.tag = `stylistic:comment-only:${this.suggestion.id}`;
        cc.appearance = "Hidden";
        cc.cannotDelete = false;
        await context.sync();

        console.log(
          `✅ [ApplySuggestionCommand] "${this.id}": comment-only insertado exitosamente`,
        );

        return { success: true, commandId: this.id };
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(`⚠️ [ApplySuggestionCommand] "${this.id}": ${message}`);
      return { success: false, commandId: this.id, error: message };
    }
  }
}
