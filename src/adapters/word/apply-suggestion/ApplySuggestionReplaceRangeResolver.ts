/* global Word, console */

import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import { applySuggestionObservability } from "../../observability/ConsoleApplySuggestionObservabilityAdapter";
import type { ApplySuggestionRangeCandidateDiagnostics } from "../../observability/ConsoleApplySuggestionObservabilityAdapter.types";
import type { TextLocator, WordSearchContainer } from "../WordTextLocatorContext.types";
import type { ApplySuggestionReviewedText } from "./ApplySuggestionReplaceRangeResolver.types";

/**
 * Re-locates the inserted/current side for replace suggestion annotation.
 */
export class ApplySuggestionReplaceRangeResolver {
  constructor(
    private readonly suggestion: Suggestion,
    private readonly textLocator: TextLocator,
    private readonly commandId: string
  ) {}

  /** Reads current/original reviewed text for one Word range. */
  async readReviewedText(
    context: Word.RequestContext,
    range: Word.Range
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
    expectedCurrentText: string
  ): Promise<boolean> {
    const reviewedText = await this.readReviewedText(context, candidate);

    return reviewedText.current === expectedCurrentText && reviewedText.original.length === 0;
  }

  /**
   * Captures the exact range shape Word exposes for one resolver candidate.
   *
   * These diagnostics are intentionally read-only. They exist to distinguish
   * whether production failures come from the mutation range itself, wrapper
   * fallback, or paragraph fallback selecting a non-current-only range.
   */
  private async inspectCandidate(
    context: Word.RequestContext,
    label: string,
    candidate: Word.Range,
    expectedCurrentText: string
  ): Promise<ApplySuggestionRangeCandidateDiagnostics> {
    candidate.load("text");
    const reviewedText = await this.readReviewedText(context, candidate);
    await context.sync();

    const diagnostics = {
      text: candidate.text,
      current: reviewedText.current,
      original: reviewedText.original,
      passes: reviewedText.current === expectedCurrentText && reviewedText.original.length === 0,
    };

    applySuggestionObservability.logResolverCandidate(this.commandId, label, diagnostics);

    return diagnostics;
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
    wrapperRange: Word.Range
  ): Promise<Word.Range | null> {
    const expectedCurrentText = this.suggestion.suggestedText ?? "";
    const mutationDiagnostics = await this.inspectCandidate(
      context,
      "mutationRange",
      mutationRange,
      expectedCurrentText
    );

    if (mutationDiagnostics.passes) {
      applySuggestionObservability.logResolverSelected(this.commandId, "mutationRange");
      return mutationRange;
    }

    const directCandidate = await this.textLocator.locate({
      context,
      container: wrapperRange as unknown as WordSearchContainer,
      searchText: expectedCurrentText,
    });

    if (!directCandidate) {
      applySuggestionObservability.logResolverCandidateNotFound(this.commandId, "wrapperRange");
    }

    if (directCandidate) {
      const directDiagnostics = await this.inspectCandidate(
        context,
        "wrapperRange candidate",
        directCandidate,
        expectedCurrentText
      );

      if (directDiagnostics.passes) {
        applySuggestionObservability.logResolverSelected(this.commandId, "wrapperRange candidate");
        return directCandidate;
      }
    }

    const paragraphRange = mutationRange.paragraphs.getFirst().getRange("Whole");
    const paragraphCandidate = await this.textLocator.locate({
      context,
      container: paragraphRange as unknown as WordSearchContainer,
      searchText: expectedCurrentText,
    });

    if (!paragraphCandidate) {
      applySuggestionObservability.logResolverCandidateNotFound(this.commandId, "paragraphRange");
    }

    if (paragraphCandidate) {
      const paragraphDiagnostics = await this.inspectCandidate(
        context,
        "paragraphRange candidate",
        paragraphCandidate,
        expectedCurrentText
      );

      if (paragraphDiagnostics.passes) {
        applySuggestionObservability.logResolverSelected(
          this.commandId,
          "paragraphRange candidate"
        );
        return paragraphCandidate;
      }
    }

    console.warn(
      `⚠️ [ApplySuggestionCommand] "${this.commandId}": no se pudo aislar el rango insertado actual (current="${mutationDiagnostics.current}", original="${mutationDiagnostics.original}")`
    );

    return null;
  }
}
