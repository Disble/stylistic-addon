/**
 * Shared TypeScript interfaces for the Stylistic add-in.
 *
 * These contracts are consumed by both the Word API layer ({@link wordApi})
 * and the business-logic layer ({@link analyzer} — Bullet 3+).
 * They intentionally carry no dependency on Office.js so that analyzers
 * can be developed and unit-tested in isolation.
 */

/** A single editorial suggestion produced by an analyzer. */
export interface Suggestion {
  /** Unique identifier (used for traceability in results). */
  id: string;

  /** Exact text to locate in the document (case-sensitive search). */
  originalText: string;

  /** Replacement text that will appear as a tracked change. */
  suggestedText: string;

  /** Human-readable reason for the suggestion. */
  justification: string;

  /**
   * When `body.search()` returns multiple matches, this zero-based index
   * selects which occurrence to replace. Defaults to 0 (first match).
   */
  paragraphIndex?: number;
}

/** Result of attempting to insert suggestions as tracked changes. */
export interface InsertionResult {
  /** Number of suggestions successfully applied. */
  successCount: number;

  /** Suggestions whose `originalText` was not found in the document. */
  failedSuggestions: Suggestion[];
}
