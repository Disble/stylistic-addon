/* global console */

/**
 * DeduplicateHandler — Phase 5 of the analysis pipeline.
 *
 * Removes suggestions with duplicate `originalText` values that arise when
 * the same phrase appears in multiple chunks. Only the first occurrence is
 * kept (case-insensitive comparison).
 *
 * Deduplication is a data integrity step: applying the same `originalText`
 * twice would search for text that is no longer present after the first
 * tracked change (since the text is wrapped in `<w:del>`), causing the
 * second suggestion to fail.
 *
 * @module DeduplicateHandler
 */

import { Suggestion } from "../../types";
import { PipelineContext } from "../PipelineContext";
import { PipelineHandler } from "./ReadTextHandler";

export class DeduplicateHandler implements PipelineHandler {
  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    const raw = ctx.rawSuggestions!;
    console.log(`🧹 [DeduplicateHandler] Fase 5: Deduplicando ${raw.length} sugerencias...`);

    const unique = this.deduplicateByOriginalText(raw);
    const removed = raw.length - unique.length;
    if (removed > 0) {
      console.log(
        `🧹 [DeduplicateHandler] ${removed} duplicado(s) removidos → ${unique.length} únicas`
      );
    }

    ctx.uniqueSuggestions = unique;
    await next();
  }

  private deduplicateByOriginalText(suggestions: Suggestion[]): Suggestion[] {
    const seen = new Set<string>();
    return suggestions.filter((s) => {
      // comment-only suggestions are never deduplicated against each other:
      // each one is unique per backend response (keyed by id) and targeting
      // the same originalText with a comment is valid across analysis runs.
      const key = s.type === "comment-only" ? s.id : s.originalText.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }
}
