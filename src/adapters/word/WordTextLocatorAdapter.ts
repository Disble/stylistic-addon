/**
 * WordTextLocatorAdapter — bridges pure text-search heuristics to Office.js
 * search execution.
 *
 * This adapter owns the operational details of locating text ranges in a Word
 * container (`Word.Body` or `Word.Range`) while delegating normalization and
 * candidate derivation to `TextSearchCore`.
 *
 * @module WordTextLocatorAdapter
 */

import {
  findFirstAlphanumericOffset,
  findUniqueLocatorSubstring,
  findWhitespaceInsensitiveSlice,
} from "../../core/text-search/TextSearchCore";
import type {
  TextLocator,
  WordSearchContainer,
  WordTextLocationRequest,
} from "./WordTextLocatorContext";

/** Result contract for one Word text-location request. */
type WordTextLocationResult = Word.Range | null;

/** Converts unknown error values into a stable string. */
function stringifyUnknownError(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  if (typeof error === "string") {
    return error;
  }

  try {
    return JSON.stringify(error);
  } catch {
    return Object.prototype.toString.call(error);
  }
}

export class WordTextLocatorAdapter implements TextLocator {
  /** Detects Word search rejections caused by invalid or too-long strings. */
  private static isSearchInvalidError(error: unknown): boolean {
    return stringifyUnknownError(error).includes(
      "SearchStringInvalidOrTooLong",
    );
  }

  /** Runs one Word search attempt and normalizes invalid-search failures. */
  private async runSearchAttempt(
    requestContext: Word.RequestContext,
    container: WordSearchContainer,
    searchText: string,
    options: Record<string, boolean>,
  ): Promise<{ invalid: boolean; results?: Word.RangeCollection }> {
    try {
      const results = container.search(searchText, options);
      results.load("items");
      await requestContext.sync();
      return { invalid: false, results };
    } catch (error) {
      if (WordTextLocatorAdapter.isSearchInvalidError(error)) {
        return { invalid: true };
      }

      throw error;
    }
  }

  /** Resolves a range through the pure fallback scan strategy. */
  private async resolveByWhitespaceScan(
    requestContext: Word.RequestContext,
    container: WordSearchContainer,
    searchText: string,
    searchOptions: Record<string, boolean>,
  ): Promise<WordTextLocationResult> {
    container.load("text");
    await requestContext.sync();

    const rawSlice = findWhitespaceInsensitiveSlice(searchText, container.text);
    if (!rawSlice) {
      return null;
    }

    const fallbackSearchText = findUniqueLocatorSubstring(
      rawSlice,
      container.text,
    );
    if (!fallbackSearchText) {
      return null;
    }

    const fallbackAttempt = await this.runSearchAttempt(
      requestContext,
      container,
      fallbackSearchText,
      searchOptions,
    );
    if (!fallbackAttempt.invalid) {
      return fallbackAttempt.results?.items[0] ?? null;
    }

    const alphanumericOffset = findFirstAlphanumericOffset(rawSlice);
    if (alphanumericOffset <= 0) {
      return null;
    }

    const alphanumericSlice = rawSlice.slice(alphanumericOffset);
    const alphanumericCandidate = findUniqueLocatorSubstring(
      alphanumericSlice,
      container.text,
    );
    if (!alphanumericCandidate) {
      return null;
    }

    const retryAttempt = await this.runSearchAttempt(
      requestContext,
      container,
      alphanumericCandidate,
      searchOptions,
    );
    if (retryAttempt.invalid) {
      return null;
    }

    return retryAttempt.results?.items[0] ?? null;
  }

  /**
   * Locates text inside a Word search container using the approved three-step
   * strategy: exact search, relaxed search, then pure-core fallback scanning.
   */
  async locate({
    context,
    container,
    searchText,
  }: WordTextLocationRequest): Promise<WordTextLocationResult> {
    const searchOptions = { matchCase: true, matchWholeWord: false };
    const relaxedOptions = {
      matchCase: true,
      matchWholeWord: false,
      ignorePunct: true,
      ignoreSpace: true,
    };

    const exactMatchAllowed = searchText.length <= 256;
    let exactResults: Word.RangeCollection | undefined;

    if (exactMatchAllowed) {
      const exactAttempt = await this.runSearchAttempt(
        context,
        container,
        searchText,
        searchOptions,
      );
      if (exactAttempt.invalid) {
        return this.resolveByWhitespaceScan(
          context,
          container,
          searchText,
          searchOptions,
        );
      }

      exactResults = exactAttempt.results;
      if ((exactResults?.items.length ?? 0) > 0) {
        return exactResults?.items[0] ?? null;
      }
    }

    const shouldRunRelaxedSearch =
      !exactMatchAllowed || (exactResults?.items.length ?? 0) === 0;
    if (shouldRunRelaxedSearch) {
      const relaxedAttempt = await this.runSearchAttempt(
        context,
        container,
        searchText,
        relaxedOptions,
      );
      if (relaxedAttempt.invalid) {
        return this.resolveByWhitespaceScan(
          context,
          container,
          searchText,
          searchOptions,
        );
      }

      if ((relaxedAttempt.results?.items.length ?? 0) > 0) {
        return relaxedAttempt.results?.items[0] ?? null;
      }
    }

    return this.resolveByWhitespaceScan(
      context,
      container,
      searchText,
      searchOptions,
    );
  }
}
