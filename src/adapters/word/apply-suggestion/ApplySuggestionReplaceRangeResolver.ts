/* global Word, console */

import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import type {
  TextLocator,
  WordSearchContainer,
} from "../WordTextLocatorContext";

/** Reviewed text pair read from one Word range. */
interface ApplySuggestionReviewedText {
  /** Current visible reviewed text. */
  current: string;

  /** Original reviewed text still associated with the range. */
  original: string;
}

/**
 * Re-locates the inserted/current side for replace suggestion annotation.
 */
export class ApplySuggestionReplaceRangeResolver {
  constructor(
    private readonly suggestion: Suggestion,
    private readonly textLocator: TextLocator,
    private readonly commandId: string,
  ) {}

  /** Reads current/original reviewed text for one Word range. */
  async readReviewedText(
    context: Word.RequestContext,
    range: Word.Range,
  ): Promise<ApplySuggestionReviewedText> {
    const current = range.getReviewedText("Current");
    const original = range.getReviewedText("Original");
    await context.sync();

    return {
      current: current.value,
      original: original.value,
    };
  }

  /**
   * Verifies whether a candidate range represents only the current inserted side.
   */
  async isCurrentOnlyReviewedRange(
    context: Word.RequestContext,
    candidate: Word.Range,
    expectedCurrentText: string,
  ): Promise<boolean> {
    const reviewedText = await this.readReviewedText(context, candidate);

    return (
      reviewedText.current === expectedCurrentText &&
      reviewedText.original.length === 0
    );
  }

  /**
   * Re-locates the current inserted side for replace suggestions.
   *
   * Office.js only guarantees that `insertText(..., replace)` returns a `Range`;
   * it does NOT guarantee that the returned range is already isolated to the
   * inserted/current side while Track Changes is on.
   */
  async resolveReplaceAnnotationRange(
    context: Word.RequestContext,
    mutationRange: Word.Range,
    wrapperRange: Word.Range,
  ): Promise<Word.Range | null> {
    const expectedCurrentText = this.suggestion.suggestedText ?? "";
    const reviewedMutationRange = await this.readReviewedText(
      context,
      mutationRange,
    );

    if (
      reviewedMutationRange.current === expectedCurrentText &&
      reviewedMutationRange.original.length === 0
    ) {
      return mutationRange;
    }

    const directCandidate = await this.textLocator.locate({
      context,
      container: wrapperRange as unknown as WordSearchContainer,
      searchText: expectedCurrentText,
    });

    if (
      directCandidate &&
      (await this.isCurrentOnlyReviewedRange(
        context,
        directCandidate,
        expectedCurrentText,
      ))
    ) {
      return directCandidate;
    }

    const paragraphRange = mutationRange.paragraphs
      .getFirst()
      .getRange("Whole");
    const paragraphCandidate = await this.textLocator.locate({
      context,
      container: paragraphRange as unknown as WordSearchContainer,
      searchText: expectedCurrentText,
    });

    if (
      paragraphCandidate &&
      (await this.isCurrentOnlyReviewedRange(
        context,
        paragraphCandidate,
        expectedCurrentText,
      ))
    ) {
      return paragraphCandidate;
    }

    console.warn(
      `⚠️ [ApplySuggestionCommand] "${this.commandId}": no se pudo aislar el rango insertado actual (current="${reviewedMutationRange.current}", original="${reviewedMutationRange.original}")`,
    );

    return null;
  }
}
