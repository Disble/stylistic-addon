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

import { IDocumentPort } from "../../domain/ports";
import { TextSource, Suggestion, InsertionResult, ProgressCallback } from "../../domain/types";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import { cleanupResolvedComments } from "./cleanup/CommentCleanup";

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
      await context.sync();

      const selText = selection.text;
      if (selText && selText.trim().length > 0) {
        console.log(`📖 [WordAdapter] Selección activa — ${selText.length} chars`);
        return { text: selText, isSelection: true };
      }

      const body = context.document.body;
      body.load("text");
      await context.sync();
      console.log(`📖 [WordAdapter] Documento completo — ${body.text.length} chars`);
      return { text: body.text, isSelection: false };
    });
  }

  /**
   * Returns the set of original texts already applied as Stylistic tracked
   * deletions. Used as a guard to prevent duplicate tracked changes on re-run.
   */
  async getAppliedOriginalTexts(): Promise<Set<string>> {
    console.log("🛡️ [WordAdapter] Consultando tracked changes de Stylistic existentes...");
    return Word.run(async (context) => {
      const tracked = context.document.body.getTrackedChanges();
      tracked.load({ select: "author,type" });
      await context.sync();

      const stylisticDeletions = tracked.items.filter(
        (tc) => tc.author === "Stylistic" && (tc.type as string) === "Deleted"
      );

      if (stylisticDeletions.length === 0) {
        return new Set<string>();
      }

      const ranges = stylisticDeletions.map((tc) => tc.getRange());
      ranges.forEach((r) => r.load("text"));
      await context.sync();

      const texts = new Set(ranges.map((r) => r.text));
      console.log(`🛡️ [WordAdapter] ${texts.size} texto(s) ya rastreado(s)`);
      return texts;
    });
  }

  /**
   * Applies suggestions as tracked changes using `ApplySuggestionCommand`
   * (Command pattern). Each suggestion runs in its own `Word.run` context
   * (per-suggestion isolation) to avoid stale ranges after OOXML insertions.
   * Command failures are aggregated as failed suggestions so later suggestions
   * can still run.
   */
  async applySuggestions(
    suggestions: Suggestion[],
    onProgress?: ProgressCallback
  ): Promise<InsertionResult> {
    console.log(`📝 [WordAdapter] applySuggestions: ${suggestions.length} sugerencias`);

    if (suggestions.length === 0) {
      return { successCount: 0, failedSuggestions: [] };
    }

    const failedSuggestions: Suggestion[] = [];
    let successCount = 0;

    for (const suggestion of suggestions) {
      const command = new ApplySuggestionCommand(suggestion);
      let commandResult;

      try {
        commandResult = await command.execute();
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        commandResult = { success: false, commandId: suggestion.id, error: message };
      }

      if (commandResult.success) {
        successCount++;
        console.log(`✅ [WordAdapter] "${suggestion.id}" aplicada`);
      } else {
        failedSuggestions.push(suggestion);
        console.warn(`⚠️ [WordAdapter] "${suggestion.id}" falló: ${commandResult.error}`);
      }

      if (onProgress) {
        onProgress(
          "applying",
          successCount + failedSuggestions.length,
          suggestions.length,
          `Aplicando sugerencia ${successCount + failedSuggestions.length} de ${suggestions.length}...`
        );
      }
    }

    console.log(
      `📝 [WordAdapter] Completado: ${successCount} éxitos, ${failedSuggestions.length} fallos`
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
}
