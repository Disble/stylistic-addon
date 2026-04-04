/**
 * WordSearchAdapter — Text-search utilities for the Word adapter layer.
 *
 * This module isolates the quirks of `Word.search()` from the command layer:
 *
 * - `Word.search()` rejects strings longer than 256 characters.
 * - Word's `body.text` may use typographic (smart) quotes and field-code control
 *   characters that differ from the backend-provided strings.
 * - Whitespace (spaces, newlines) in the document may not match the backend text.
 *
 * All functions here are pure (no Office.js calls) and independently testable.
 *
 * @module WordSearchAdapter
 */

type IndexedText = {
  text: string;
  indices: number[];
};

/** Maximum string length accepted by `Word.search()`. */
const WORD_SEARCH_MAX_LENGTH = 256;

/**
 * Normalizes a single character for cross-source comparison.
 *
 * Word's `body.text` may differ from the backend-provided string in several ways:
 * 1. Typographic/smart quotes (`\u201C`, `\u201D`, `\u2018`, `\u2019`) vs. straight quotes.
 * 2. Word control characters (`\u0013`–`\u0015`, field delimiters).
 * 3. Diacritic accents: the backend may send the anchor with the suggested
 *    diacritic already applied (e.g. `qué`) while the document still has the
 *    uncorrected form (`que`), or vice versa. We normalize by stripping diacritics
 *    so both forms compare equal during the whitespace-insensitive search step.
 *
 * Returns a canonical replacement, or the original character unchanged.
 */
export function normalizeChar(char: string): string {
  // Smart double quotes → straight double quote
  if (char === "\u201C" || char === "\u201D") return '"';
  // Smart single quotes / apostrophes → straight apostrophe
  if (char === "\u2018" || char === "\u2019") return "'";
  // Word field-code control characters → empty (invisible, no semantic content)
  if (char >= "\u0013" && char <= "\u0015") return "";
  // Strip diacritics: decompose to NFD (base + combining mark), then remove
  // combining diacritical marks (Unicode block U+0300–U+036F). This lets
  // "qué" match "que" and "así" match "asi" during anchor search.
  const decomposed = char.normalize("NFD");
  return decomposed.replace(/[\u0300-\u036f]/g, "");
}

/**
 * Strips whitespace and normalizes cross-source characters, tracking each
 * surviving character's original index for slice reconstruction.
 *
 * Word field codes are delimited by `\u0013` (field start) and `\u0015`
 * (field end). The content between these delimiters is not user-visible
 * semantic text, so the entire field — delimiters and content — is skipped.
 *
 * Used by `findWhitespaceInsensitiveSlice` to align text from the backend
 * with text returned by Word's `body.text`.
 */
export function removeWhitespaceWithIndices(text: string): IndexedText {
  const indices: number[] = [];
  let normalized = "";
  let insideFieldCode = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];

    // Word field-code start delimiter — skip until matching end delimiter
    if (char === "\u0013") {
      insideFieldCode = true;
      continue;
    }
    // Word field-code end delimiter — resume normal scanning
    if (char === "\u0015") {
      insideFieldCode = false;
      continue;
    }
    if (insideFieldCode) {
      continue;
    }

    if (/\s/.test(char)) {
      continue;
    }

    const replaced = normalizeChar(char);
    if (replaced === "") {
      continue;
    }

    normalized += replaced;
    indices.push(index);
  }

  return { text: normalized, indices };
}

/**
 * Locates `searchText` within `documentText` using whitespace-insensitive,
 * quote-normalized comparison.
 *
 * Returns the exact slice of `documentText` (preserving original spacing and
 * characters) that corresponds to the match, or `null` if not found.
 *
 * This is the fallback used when `container.search()` returns no results or
 * throws `SearchStringInvalidOrTooLong`.
 */
export function findWhitespaceInsensitiveSlice(
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
 * Returns a string safe for use with `Word.search()` that uniquely identifies
 * `slice` within `containerText`.
 *
 * - If `slice.length <= 256`: returns the slice as-is — it already fits within
 *   Word's search limit, no truncation needed.
 * - If `slice.length > 256`: searches for the shortest prefix of `slice`
 *   (starting from length 1, up to 256) that appears exactly once in
 *   `containerText`. Returns `null` if no unique prefix exists within the limit.
 *
 * This guarantees that the string passed to `Word.search()` both fits within
 * the API's 256-char limit AND uniquely identifies the target range, avoiding
 * false matches when the same prefix repeats elsewhere in the document.
 */
export function findUniqueLocatorSubstring(
  slice: string,
  containerText: string,
): string | null {
  // Slice fits within Word's limit — return as-is, no truncation needed.
  if (slice.length <= WORD_SEARCH_MAX_LENGTH) {
    return slice;
  }

  // Slice exceeds the limit — find the shortest unique prefix within 256 chars.
  for (let length = 1; length <= WORD_SEARCH_MAX_LENGTH; length += 1) {
    const candidate = slice.slice(0, length);
    const firstIndex = containerText.indexOf(candidate);
    if (firstIndex === -1) {
      // Should not happen if slice came from containerText, but guard anyway
      return null;
    }
    const isUnique = !containerText.includes(candidate, firstIndex + 1);
    if (isUnique) {
      return candidate;
    }
  }

  return null;
}

/**
 * Returns the index of the first alphanumeric character in `text`, or -1 if
 * none is found.
 *
 * Used as a fallback when `Word.search()` rejects a candidate that starts with
 * special characters (em-dashes, inverted question marks, typographic quotes,
 * etc.). Skipping to the first letter/digit produces a candidate that Word can
 * accept without throwing `SearchStringInvalidOrTooLong`.
 */
export function findFirstAlphanumericOffset(text: string): number {
  for (let i = 0; i < text.length; i += 1) {
    if (/[a-zA-Z0-9\u00C0-\u024F]/.test(text[i])) {
      return i;
    }
  }
  return -1;
}
