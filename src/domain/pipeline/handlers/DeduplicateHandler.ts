/* global console */

/**
 * DeduplicateHandler — Phase 5 of the analysis pipeline.
 *
 * Removes suggestions with duplicate `(context, anchor)` pairs that arise when
 * the same suggestion is returned more than once across chunks. Only the first
 * occurrence is kept (case-insensitive comparison).
 *
 * Deduplication is a data integrity step: applying the same anchor inside the
 * same context twice would search for text that is no longer present after the
 * first tracked change (since the text is wrapped in `<w:del>`), causing the
 * second suggestion to fail.
 *
 * @module DeduplicateHandler
 */

import type { Suggestion } from "../../suggestion/Suggestion.types";
import type { PipelineContext } from "../PipelineContext";
import type { PipelineHandler } from "./ReadTextHandler";

export class DeduplicateHandler implements PipelineHandler {
  async handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void> {
    const raw = ctx.rawSuggestions!;
    console.log(
      `🧹 [DeduplicateHandler] Fase 5: Deduplicando ${raw.length} sugerencias...`,
    );

    const unique = this.deduplicateByContextAnchor(raw);
    const removed = raw.length - unique.length;
    if (removed > 0) {
      console.log(
        `🧹 [DeduplicateHandler] ${removed} duplicado(s) removidos → ${unique.length} únicas`,
      );
    }

    ctx.uniqueSuggestions = unique;
    await next();
  }

  /** Deduplicates track-change suggestions by a composite `context:anchor` key. */
  private deduplicateByContextAnchor(suggestions: Suggestion[]): Suggestion[] {
    const seen = new Set<string>();
    return suggestions.filter((s) => {
      // comment-only suggestions are never deduplicated against each other:
      // each one is unique per backend response (keyed by id) and targeting
      // the same anchor with a comment is valid across analysis runs.
      const key =
        s.type === "comment-only"
          ? s.id
          : `${s.context}:${s.anchor}`.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }
}
