/* global console */

/**
 * DeduplicateHandler — Phase 5 of the analysis pipeline.
 *
 * Removes duplicate suggestions that arise when the same semantic issue is
 * returned more than once across chunks. Track-change suggestions use their
 * mutating `(context, anchor)` target; comment-only suggestions use the exact
 * visible comment identity, not their transport id.
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

  /** Deduplicates suggestions by the semantic identity that would hit Word. */
  private deduplicateByContextAnchor(suggestions: Suggestion[]): Suggestion[] {
    const seen = new Set<string>();
    return suggestions.filter((s) => {
      const key = this.buildDeduplicationKey(s);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  /** Builds the deduplication key without treating transport ids as identity. */
  private buildDeduplicationKey(suggestion: Suggestion): string {
    if (suggestion.type === "comment-only") {
      return [
        suggestion.type,
        suggestion.context,
        suggestion.anchor,
        suggestion.justification,
        suggestion.category,
        suggestion.severity,
      ]
        .map((part) => this.normalizeKeyPart(part))
        .join(":");
    }

    return `${suggestion.context}:${suggestion.anchor}`.toLowerCase();
  }

  /** Normalizes free-text fields so cosmetic casing/spacing cannot bypass dedupe. */
  private normalizeKeyPart(part: string | undefined): string {
    return (part ?? "").trim().replace(/\s+/g, " ").toLowerCase();
  }
}
