/* global Word, console */

/**
 * Word Adapter — implements `IDocumentPort` using Office.js.
 *
 * This is the ONLY file in the codebase that references the `Word` global.
 * All other modules interact with Word documents exclusively through this
 * adapter via the `IDocumentPort` interface.
 *
 * Responsibilities:
 * - Read document text (full body or active selection).
 * - Query existing Stylistic tracked changes (Guard pattern).
 * - Apply suggestions as tracked changes via `ApplySuggestionCommand`.
 * - Delegate comment cleanup to `CommentCleanup`.
 *
 * @module WordAdapter
 */

import type { IDocumentPort } from "../../domain/ports";
import type {
  CommandResult,
  InsertionResult,
  ProgressCallback,
  Suggestion,
  SuggestionActionResult,
  TextSource,
} from "../../domain/types";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import {
  COMMENT_ONLY_TAG_PREFIX,
  cleanupResolvedComments,
  OVERLAPPING_RELATIONS,
} from "./cleanup/CommentCleanup";

type ParagraphSnapshot = {
  text?: string;
  styleBuiltIn?: string;
  firstLineIndent?: number;
  leftIndent?: number;
};

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

export class WordAdapter implements IDocumentPort {
  /**
   * Resolves the text to analyze: returns the current selection if non-empty,
   * otherwise falls back to the full document body (Transparent Fallback pattern).
   */
  async getTextToAnalyze(): Promise<TextSource> {
    console.log("📖 [WordAdapter] Resolviendo texto a analizar...");
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      selection.paragraphs.load(PARAGRAPH_LOAD_FIELDS);
      await context.sync();

      const hasSelectedText = selection.text.trim().length > 0;
      const selText = hasSelectedText
        ? buildStructuredParagraphText(
            selection.paragraphs.items as ParagraphSnapshot[],
          ) || selection.text
        : "";

      if (selText && selText.trim().length > 0) {
        console.log(
          `📖 [WordAdapter] Selección activa — ${selText.length} chars`,
        );
        return { text: selText, isSelection: true };
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

  /**
   * Returns the set of original texts already applied as Stylistic tracked
   * deletions OR as active comment-only suggestions.
   *
   * Used as a guard to prevent duplicate tracked changes and duplicate
   * comment-only suggestions on re-run.
   *
   * For track-change suggestions: reads the deleted range text from the TCs.
   * For comment-only suggestions: reads the range text of the CC anchor itself,
   * which spans the original text at insertion time.
   */
  async getAppliedOriginalTexts(): Promise<Set<string>> {
    console.log(
      "🛡️ [WordAdapter] Consultando tracked changes y CCs comment-only de Stylistic...",
    );
    return Word.run(async (context) => {
      const tracked = context.document.body.getTrackedChanges();
      const allCCs = context.document.contentControls;
      tracked.load({ select: "author,type" });
      allCCs.load({ select: "tag" });
      await context.sync();

      const stylisticDeletions = tracked.items.filter(
        (tc) => tc.author === "Stylistic" && (tc.type as string) === "Deleted",
      );

      // JS-side prefix filter — Office.js getByTag() is exact-match only
      const commentOnlyCCs = allCCs.items.filter((cc) =>
        cc.tag.startsWith(COMMENT_ONLY_TAG_PREFIX),
      );

      if (stylisticDeletions.length === 0 && commentOnlyCCs.length === 0) {
        return new Set<string>();
      }

      const tcRanges = stylisticDeletions.map((tc) => tc.getRange());
      tcRanges.forEach((r) => {
        r.load("text");
      });

      const ccRanges = commentOnlyCCs.map((cc) => cc.getRange());
      ccRanges.forEach((r) => {
        r.load("text");
      });

      await context.sync();

      const texts = new Set([
        ...tcRanges.map((r) => r.text),
        ...ccRanges.map((r) => r.text),
      ]);
      console.log(
        `🛡️ [WordAdapter] ${texts.size} texto(s) ya rastreado(s) (TC + comment-only)`,
      );
      return texts;
    });
  }

  /**
   * Sorts suggestions in reverse array order as a heuristic for applying
   * end-of-document suggestions first.
   *
   * The backend returns suggestions in reading order (start → end of document).
   * Reversing ensures each applied Content Control only affects text BEFORE the
   * next search target, preventing CC boundary interference when multiple
   * suggestions target the same paragraph.
   *
   * This is a heuristic: it assumes backend output order approximates document
   * order. A future improvement could sort by actual `Range` position using
   * `Range.compareLocationWith()` inside a single `Word.run`.
   */
  private sortByDocumentPosition(suggestions: Suggestion[]): Suggestion[] {
    if (suggestions.length <= 1) return suggestions;
    return [...suggestions].reverse();
  }

  /**
   * Applies suggestions as tracked changes using `ApplySuggestionCommand`
   * (Command pattern). Each suggestion runs in its own `Word.run` context
   * (per-suggestion isolation) to avoid stale ranges after OOXML insertions.
   * Suggestions are applied in reverse order (end-of-document first) to avoid
   * Content Control boundary interference when multiple suggestions share the
   * same paragraph. Command failures are aggregated as failed suggestions so
   * later suggestions can still run.
   */
  async applySuggestions(
    suggestions: Suggestion[],
    onProgress?: ProgressCallback,
  ): Promise<InsertionResult> {
    console.log(
      `📝 [WordAdapter] applySuggestions: ${suggestions.length} sugerencias`,
    );

    if (suggestions.length === 0) {
      return { successCount: 0, failedSuggestions: [] };
    }

    // Apply end-of-document suggestions first so earlier searches are unaffected
    // by Content Controls created for later positions in the same paragraph.
    const sortedSuggestions = this.sortByDocumentPosition(suggestions);

    const failedSuggestions: Suggestion[] = [];
    let successCount = 0;

    for (const suggestion of sortedSuggestions) {
      const command = new ApplySuggestionCommand(suggestion);
      let commandResult: CommandResult;

      try {
        commandResult = await command.execute();
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        commandResult = {
          success: false,
          commandId: suggestion.id,
          error: message,
        };
      }

      if (commandResult.success) {
        successCount++;
        console.log(`✅ [WordAdapter] "${suggestion.id}" aplicada`);
      } else {
        failedSuggestions.push(suggestion);
        console.warn(
          `⚠️ [WordAdapter] "${suggestion.id}" falló: ${commandResult.error}`,
        );
      }

      if (onProgress) {
        onProgress(
          "applying",
          successCount + failedSuggestions.length,
          suggestions.length,
          `Aplicando sugerencia ${successCount + failedSuggestions.length} de ${suggestions.length}...`,
        );
      }
    }

    console.log(
      `📝 [WordAdapter] Completado: ${successCount} éxitos, ${failedSuggestions.length} fallos`,
    );
    return { successCount, failedSuggestions };
  }

  /**
   * Deletes Stylistic comments whose tracked changes have been resolved.
   * Delegates to the `CommentCleanup` module (Range Colocation pattern).
   */
  async cleanupResolvedComments(): Promise<{ deleted: number; kept: number }> {
    return cleanupResolvedComments();
  }

  /**
   * Finds the Content Control for a suggestion, supporting both the new tag
   * format (`stylistic:{type}:{id}`) and the legacy bare-ID format.
   *
   * Office.js `getByTag()` is exact-match only, so we try both formats and
   * return the first CC found. Returns `null` if not found under either format.
   */
  private findCCByTag(
    context: Word.RequestContext,
    suggestion: Suggestion,
  ): {
    newFormatResult: Word.ContentControlCollection;
    legacyResult: Word.ContentControlCollection;
  } {
    const newTag = `stylistic:${suggestion.type}:${suggestion.id}`;
    const legacyTag = suggestion.id;
    const newFormatResult = context.document.contentControls.getByTag(newTag);
    const legacyResult = context.document.contentControls.getByTag(legacyTag);
    newFormatResult.load("items");
    legacyResult.load("items");
    return { newFormatResult, legacyResult };
  }

  /**
   * Shared implementation for accepting or rejecting a suggestion.
   *
   * Branches on `suggestion.type`:
   * - `"track-change"`: finds Stylistic TCs inside the CC, accepts/rejects them,
   *   deletes the colocated comment, then deletes the CC.
   * - `"comment-only"`: skips TC lookup entirely. Finds and deletes the Word
   *   comment anchored to the CC range, then deletes the CC.
   *
   * Also supports legacy CCs tagged with a bare suggestion ID (no `stylistic:`
   * prefix) — those are treated as `"track-change"` for backward compatibility.
   *
   * Never throws — catches all errors and returns a result object.
   */
  private async resolveSuggestion(
    suggestion: Suggestion,
    action: "accept" | "reject",
  ): Promise<SuggestionActionResult> {
    try {
      return await Word.run(async (context) => {
        // 1. Find the Content Control — try new tag format first, fall back to legacy bare ID
        const { newFormatResult, legacyResult } = this.findCCByTag(
          context,
          suggestion,
        );
        await context.sync();

        let cc: Word.ContentControl | null = null;
        if (newFormatResult.items.length > 0) {
          cc = newFormatResult.items[0];
        } else if (legacyResult.items.length > 0) {
          cc = legacyResult.items[0];
          console.log(
            `🔁 [WordAdapter] "${suggestion.id}": usando tag legado (bare ID)`,
          );
        }

        if (!cc) {
          return {
            status: "cc-not-found" as const,
            trackedChangesAffected: 0,
            commentDeleted: false,
          };
        }

        // 2. Find and delete the colocated Stylistic comment (shared by both branches)
        const comments = context.document.body.getComments();
        comments.load({ select: "authorName" });
        await context.sync();

        const stylisticComments = comments.items.filter(
          (c) => c.authorName === "Stylistic",
        );
        let commentDeleted = false;
        const ccRange = cc.getRange();

        for (const comment of stylisticComments) {
          const commentRange = comment.getRange();
          const locationResult = commentRange.compareLocationWith(ccRange);
          await context.sync();
          if (OVERLAPPING_RELATIONS.includes(locationResult.value as string)) {
            comment.delete();
            commentDeleted = true;
            break;
          }
        }

        // 3a. comment-only branch: no TCs to process — just delete the CC
        if (suggestion.type === "comment-only") {
          cc.delete(true); // true = keep content
          await context.sync();
          console.log(
            `🗨️ [WordAdapter] "${suggestion.id}": comment-only ${action}ed, comentario eliminado: ${commentDeleted}`,
          );
          return {
            status:
              action === "accept"
                ? ("accepted" as const)
                : ("rejected" as const),
            trackedChangesAffected: 0,
            commentDeleted,
          };
        }

        // 3b. track-change branch: accept/reject TCs inside the CC
        // The CC tag already uniquely identifies this suggestion — all TCs
        // inside it belong to it. No author filter needed or reliable.
        const tcs = cc.getTrackedChanges();
        tcs.load("items");
        await context.sync();

        if (tcs.items.length === 0) {
          // TCs already resolved but CC still present — clean up the orphaned CC
          cc.delete(true); // true = keep content
          await context.sync();
          return {
            status: "already-resolved" as const,
            trackedChangesAffected: 0,
            commentDeleted,
          };
        }

        for (const tc of tcs.items) {
          if (action === "accept") {
            tc.accept();
          } else {
            tc.reject();
          }
        }

        // 4. Delete the Content Control anchor itself
        cc.delete(true); // true = keep content
        await context.sync();

        return {
          status:
            action === "accept" ? ("accepted" as const) : ("rejected" as const),
          trackedChangesAffected: tcs.items.length,
          commentDeleted,
        };
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return {
        status: "error" as const,
        trackedChangesAffected: 0,
        commentDeleted: false,
        error: message,
      };
    }
  }

  /**
   * Accepts all Stylistic tracked changes associated with a suggestion.
   * Also deletes the associated Stylistic comment if present.
   * Never throws — returns a `SuggestionActionResult`.
   */
  async acceptSuggestion(
    suggestion: Suggestion,
  ): Promise<SuggestionActionResult> {
    return this.resolveSuggestion(suggestion, "accept");
  }

  /**
   * Rejects all Stylistic tracked changes associated with a suggestion.
   * Also deletes the associated Stylistic comment if present.
   * Never throws — returns a `SuggestionActionResult`.
   */
  async rejectSuggestion(
    suggestion: Suggestion,
  ): Promise<SuggestionActionResult> {
    return this.resolveSuggestion(suggestion, "reject");
  }

  /**
   * Navigates the Word document to the first occurrence of `text` by selecting
   * the matching range (Word scrolls to the selection automatically).
   * Never throws — silently no-ops when text is not found or on any error.
   */
  async navigateToText(text: string): Promise<void> {
    try {
      await Word.run(async (context) => {
        const results = context.document.body.search(text, {
          matchCase: true,
          matchWholeWord: false,
        });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) {
          results.items[0].select();
          await context.sync();
        }
      });
    } catch {
      // Navigation is best-effort — silently ignore all failures
    }
  }
}
