/**
 * Text analysis engine — produces editorial {@link Suggestion}s from plain text.
 *
 * This module is the business-logic core of Stylistic. It **never** imports
 * Office.js or any DOM API, keeping it pure and unit-testable in isolation.
 *
 * Architecture:
 * - Each editorial concern is encapsulated in an {@link AnalysisRule}.
 * - Rules are built declaratively via {@link buildPatternRule}, which takes
 *   a list of {@link PatternEntry} (regex + replacement + justification).
 * - The public {@link analyze} function runs every registered rule against
 *   the input text and returns a deduplicated list of suggestions.
 *
 * Adding a new rule:
 * 1. Define a `PatternEntry[]` array with your patterns.
 * 2. Call `buildPatternRule(id, category, entries)`.
 * 3. Append the result to the `rules` array.
 *
 * @module analyzer
 */

import { Suggestion } from "./types";

// ---------------------------------------------------------------------------
// Internal interfaces
// ---------------------------------------------------------------------------

/**
 * A self-contained analysis rule that scans text and produces suggestions.
 *
 * Each rule owns a single editorial concern (e.g. redundancy, filler words).
 * Implementing new concerns means creating new `AnalysisRule` instances,
 * not modifying existing ones (Open/Closed Principle).
 */
interface AnalysisRule {
  /** Stable identifier used as prefix for suggestion IDs (e.g. "redundancy"). */
  id: string;

  /** Human-readable category label included in justification messages. */
  category: string;

  /**
   * Scans `text` and returns all detected suggestions for this rule.
   *
   * @param text - The full plain-text content of the document.
   * @returns Zero or more suggestions. Never throws.
   */
  detect(text: string): Suggestion[];
}

/**
 * Declarative definition of a single pattern within a rule.
 *
 * Each entry maps a regex match to a replacement string and an explanation.
 * Used by {@link buildPatternRule} to construct an {@link AnalysisRule}.
 */
interface PatternEntry {
  /**
   * Regex to match against the document text.
   * Use the `g` flag for global matching and `i` for case-insensitive.
   */
  pattern: RegExp;

  /** Text that should replace the match. Use `""` to suggest deletion. */
  replacement: string;

  /** Explanation of why this change improves the text (shown to the user). */
  justification: string;
}

// ---------------------------------------------------------------------------
// Rule factory
// ---------------------------------------------------------------------------

/**
 * Factory that creates an {@link AnalysisRule} from a declarative list of patterns.
 *
 * This avoids repeating the scan-and-collect loop for every rule. Each rule
 * only needs to declare *what* to find and *how* to fix it.
 *
 * @param id       - Stable rule identifier (used as suggestion ID prefix).
 * @param category - Human-readable label prepended to justification strings.
 * @param entries  - Array of pattern-replacement-justification triples.
 * @returns A fully functional {@link AnalysisRule}.
 */
function buildPatternRule(
  id: string,
  category: string,
  entries: PatternEntry[]
): AnalysisRule {
  return {
    id,
    category,
    detect(text: string): Suggestion[] {
      const suggestions: Suggestion[] = [];
      let counter = 0;

      for (const entry of entries) {
        // Create a fresh regex to reset lastIndex for each invocation
        const regex = new RegExp(entry.pattern.source, entry.pattern.flags);
        let match: RegExpExecArray | null;

        while ((match = regex.exec(text)) !== null) {
          suggestions.push({
            id: `${id}-${counter++}`,
            originalText: match[0],
            suggestedText: entry.replacement,
            justification: `[${category}] ${entry.justification}`,
          });

          // Non-global regexes would loop forever without this guard
          if (!regex.global) break;
        }
      }

      return suggestions;
    },
  };
}

// ---------------------------------------------------------------------------
// Rule definitions
// ---------------------------------------------------------------------------

/**
 * Detects pleonasms and redundant modifiers in Spanish text.
 *
 * Examples: "completamente necesario" → "necesario",
 *           "subir arriba" → "subir".
 */
const redundancyRule = buildPatternRule("redundancy", "Redundancia", [
  {
    pattern: /completamente necesario/gi,
    replacement: "necesario",
    justification: '"Necesario" ya implica completitud.',
  },
  {
    pattern: /totalmente imprescindible/gi,
    replacement: "imprescindible",
    justification: '"Imprescindible" ya es absoluto.',
  },
  {
    pattern: /muy único/gi,
    replacement: "único",
    justification: '"Único" no admite grados.',
  },
  {
    pattern: /resultado final/gi,
    replacement: "resultado",
    justification: "Un resultado es, por definición, final.",
  },
  {
    pattern: /periodo de tiempo/gi,
    replacement: "periodo",
    justification: '"Periodo" ya denota tiempo.',
  },
  {
    pattern: /volver a repetir/gi,
    replacement: "repetir",
    justification: '"Repetir" ya implica volver a hacer.',
  },
  {
    pattern: /subir arriba/gi,
    replacement: "subir",
    justification: '"Subir" ya indica dirección ascendente.',
  },
  {
    pattern: /bajar abajo/gi,
    replacement: "bajar",
    justification: '"Bajar" ya indica dirección descendente.',
  },
]);

/**
 * Detects filler words and phrases that weaken prose without adding meaning.
 *
 * Replacements are empty strings — the filler is simply removed.
 * Examples: "Básicamente, ..." → "...", "Obviamente, ..." → "...".
 */
const fillerRule = buildPatternRule("filler", "Muletilla", [
  {
    pattern: /en realidad,?\s*/gi,
    replacement: "",
    justification: "Muletilla que rara vez aporta significado.",
  },
  {
    pattern: /básicamente,?\s*/gi,
    replacement: "",
    justification: "Muletilla que debilita la afirmación.",
  },
  {
    pattern: /obviamente,?\s*/gi,
    replacement: "",
    justification: "Si es obvio, no necesita decirse.",
  },
  {
    pattern: /como ya se sabe,?\s*/gi,
    replacement: "",
    justification: "Asume conocimiento del lector innecesariamente.",
  },
]);

/**
 * Suggests simpler or more direct alternatives for verbose expressions.
 *
 * Examples: "utilizar" → "usar", "con el objetivo de" → "para".
 */
const wordChoiceRule = buildPatternRule("wordchoice", "Elección de palabra", [
  {
    pattern: /realizar/gi,
    replacement: "hacer",
    justification: '"Hacer" es más directo y claro que "realizar".',
  },
  {
    pattern: /utilizar/gi,
    replacement: "usar",
    justification: '"Usar" es más simple y directo que "utilizar".',
  },
  {
    pattern: /a nivel de/gi,
    replacement: "en",
    justification: '"En" es más conciso que "a nivel de".',
  },
  {
    pattern: /con el objetivo de/gi,
    replacement: "para",
    justification: '"Para" es más directo que "con el objetivo de".',
  },
  {
    pattern: /en el caso de que/gi,
    replacement: "si",
    justification: '"Si" es más conciso que "en el caso de que".',
  },
  {
    pattern: /debido al hecho de que/gi,
    replacement: "porque",
    justification: '"Porque" es más directo.',
  },
]);

// ---------------------------------------------------------------------------
// Rule registry
// ---------------------------------------------------------------------------

/** All active analysis rules. Add new rules here to include them in analysis. */
const rules: AnalysisRule[] = [redundancyRule, fillerRule, wordChoiceRule];

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Analyzes plain text and returns editorial suggestions.
 *
 * Runs every registered {@link AnalysisRule} against the input, collects all
 * matches, and deduplicates by `originalText` (case-insensitive) so that
 * the same phrase is only suggested once even if multiple rules flag it.
 *
 * @param text - The full document text (plain string, no Office.js dependency).
 * @returns A deduplicated array of {@link Suggestion} objects.
 *          Returns an empty array if no issues are detected.
 */
export function analyze(text: string): Suggestion[] {
  const allSuggestions: Suggestion[] = [];

  for (const rule of rules) {
    allSuggestions.push(...rule.detect(text));
  }

  return deduplicateByOriginalText(allSuggestions);
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Removes duplicate suggestions that target the same original text.
 *
 * When multiple rules flag the same phrase (e.g. a filler that is also
 * a redundancy), only the first match is kept. Comparison is case-insensitive.
 *
 * @param suggestions - The raw suggestion list (may contain duplicates).
 * @returns A filtered list with unique `originalText` values.
 */
function deduplicateByOriginalText(suggestions: Suggestion[]): Suggestion[] {
  const seen = new Set<string>();
  return suggestions.filter((s) => {
    const key = s.originalText.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
