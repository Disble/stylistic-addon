/**
 * TextSearchCore — Canonical pure text-search primitives.
 *
 * This module owns the reusable normalization and matching behavior shared by
 * Word-facing adapters. It is intentionally pure and framework-free so the
 * matching rules can be tested independently from Office.js host execution.
 *
 * @module TextSearchCore
 */

import type { IndexedText } from "./TextSearchCore.types";
import { DEFAULT_WORD_SEARCH_MAX_LENGTH } from "./TextSearchCore.constants";

/**
 * Normalizes a single character for cross-source comparison.
 *
 * The text returned by Word may differ from backend-provided text in smart
 * quotes, field-code delimiters, and diacritics. This function maps those
 * differences into a canonical comparison form.
 */
export function normalizeChar(char: string): string {
  if (char === "\u201C" || char === "\u201D") {
    return '"';
  }

  if (char === "\u2018" || char === "\u2019") {
    return "'";
  }

  if (char >= "\u0013" && char <= "\u0015") {
    return "";
  }

  const decomposed = char.normalize("NFD");
  return decomposed.replace(/[\u0300-\u036f]/g, "");
}

/**
 * Removes whitespace and non-semantic Word field-code content while preserving
 * original character indices for later slice reconstruction.
 */
export function removeWhitespaceWithIndices(text: string): IndexedText {
  const indices: number[] = [];
  let normalized = "";
  let insideFieldCode = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];

    if (char === "\u0013") {
      insideFieldCode = true;
      continue;
    }

    if (char === "\u0015") {
      insideFieldCode = false;
      continue;
    }

    if (insideFieldCode || /\s/.test(char)) {
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
 * Locates `searchText` inside `documentText` using whitespace-insensitive,
 * quote-normalized, diacritic-tolerant comparison.
 *
 * Returns the exact slice from the original `documentText`, preserving the
 * document's original spacing and typography.
 */
export function findWhitespaceInsensitiveSlice(
  searchText: string,
  documentText: string
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
  const end = normalizedDocument.indices[matchIndex + normalizedSearch.length - 1] + 1;

  return documentText.slice(start, end);
}

/**
 * Returns the shortest prefix of `slice` that still identifies the target
 * uniquely within `containerText` while respecting the Word search length cap.
 */
export function findUniqueLocatorSubstring(
  slice: string,
  containerText: string,
  maxLength = DEFAULT_WORD_SEARCH_MAX_LENGTH
): string | null {
  if (slice.length <= maxLength) {
    return slice;
  }

  for (let length = 1; length <= maxLength; length += 1) {
    const candidate = slice.slice(0, length);
    const firstIndex = containerText.indexOf(candidate);
    if (firstIndex === -1) {
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
 * Returns the first alphanumeric offset in `text`, or `-1` when none exists.
 */
export function findFirstAlphanumericOffset(text: string): number {
  for (let index = 0; index < text.length; index += 1) {
    if (/[a-zA-Z0-9\u00C0-\u024F]/.test(text[index])) {
      return index;
    }
  }

  return -1;
}
