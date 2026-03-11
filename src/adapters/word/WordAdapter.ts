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
import { TextSource, Suggestion, InsertionResult, ProgressCallback, SuggestionActionResult } from "../../domain/types";
import { ApplySuggestionCommand } from "./ApplySuggestionCommand";
import { cleanupResolvedComments, OVERLAPPING_RELATIONS } from "./cleanup/CommentCleanup";

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

  /**
   * Shared implementation for accepting or rejecting Stylistic tracked changes
   * associated with a suggestion. Also deletes the colocated Stylistic comment.
   * Never throws — catches all errors and returns a result object.
   */
  private async resolveSuggestion(
    suggestion: Suggestion,
    action: "accept" | "reject"
  ): Promise<SuggestionActionResult> {
    try {
      return await Word.run(async (context) => {
        // 1. Find the Content Control by the suggestion ID tag
        const ccs = context.document.contentControls.getByTag(suggestion.id);
        ccs.load("items");
        await context.sync();

        if (ccs.items.length === 0) {
          return {
            status: "already-resolved" as const,
            trackedChangesAffected: 0,
            commentDeleted: false,
          };
        }

        const cc = ccs.items[0];

        // 2. Get all tracked changes inside the Content Control
        const tcs = cc.getTrackedChanges();
        tcs.load("items");
        await context.sync();

        const stylisticTCs = tcs.items.filter((tc: any) => tc.author === "Stylistic");

        if (stylisticTCs.length === 0) {
          // If the TCs are gone but the CC is left behind, clean it up
          cc.delete(true); // true = keep content
          await context.sync();
          return {
            status: "already-resolved" as const,
            trackedChangesAffected: 0,
            commentDeleted: false,
          };
        }

        // 3. Accept or reject the tracked changes
        for (const tc of stylisticTCs) {
          if (action === "accept") {
            tc.accept();
          } else {
            tc.reject();
          }
        }

        // 4. Find and delete the colocated Stylistic comment
        const comments = context.document.body.getComments();
        comments.load("items");
        await context.sync();

        let commentDeleted = false;
        const ccRange = cc.getRange();

        for (const comment of comments.items) {
          if (comment.authorName !== "Stylistic") continue;
          const commentRange = comment.getRange();
          const locationResult = commentRange.compareLocationWith(ccRange);
          await context.sync();
          if (OVERLAPPING_RELATIONS.includes(locationResult.value as string)) {
            comment.delete();
            commentDeleted = true;
            break;
          }
        }

        // 5. Delete the Content Control anchor itself
        cc.delete(true); // true = keep content

        await context.sync();

        return {
          status: action === "accept" ? ("accepted" as const) : ("rejected" as const),
          trackedChangesAffected: stylisticTCs.length,
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
  async acceptSuggestion(suggestion: Suggestion): Promise<SuggestionActionResult> {
    return this.resolveSuggestion(suggestion, "accept");
  }

  /**
   * Rejects all Stylistic tracked changes associated with a suggestion.
   * Also deletes the associated Stylistic comment if present.
   * Never throws — returns a `SuggestionActionResult`.
   */
  async rejectSuggestion(suggestion: Suggestion): Promise<SuggestionActionResult> {
    return this.resolveSuggestion(suggestion, "reject");
  }
}
