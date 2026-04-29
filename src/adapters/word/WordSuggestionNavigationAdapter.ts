/* global Word */

import type {
  Suggestion,
  SuggestionNavigationResult,
  SuggestionObservationStatus,
} from "../../domain/suggestion/Suggestion.types";
import { SuggestionArtifactLocator } from "./location/SuggestionArtifactLocator";
import { SuggestionTextRangeLocator } from "./location/SuggestionTextRangeLocator";
import type {
  TextLocator,
  WordSearchContainer,
} from "./WordTextLocatorContext";

/**
 * Handles read-only suggestion navigation in Word.
 *
 * Navigation is intentionally stricter than a generic text search because a
 * wrong selection is a user-visible defect: the cursor appears to validate the
 * wrong occurrence. The adapter therefore uses the same safe localization
 * strategies as the mutating workflows, but never executes those workflows:
 *
 * 1. Prefer the persisted Word artifact for the suggestion.
 * 2. Fall back only to `context -> anchor` textual localization.
 * 3. Never search `anchor` globally when `context` is missing.
 * 4. Return a semantic no-navigation result for ambiguity or host failures.
 */
export class WordSuggestionNavigationAdapter {
  private readonly artifactLocator: SuggestionArtifactLocator;
  private readonly textRangeLocator: SuggestionTextRangeLocator;

  constructor(
    private readonly textLocator: TextLocator,
    artifactLocator = new SuggestionArtifactLocator(),
    textRangeLocator = new SuggestionTextRangeLocator(textLocator),
  ) {
    this.artifactLocator = artifactLocator;
    this.textRangeLocator = textRangeLocator;
  }

  /** Searches a body or range through the shared Word text locator. */
  async searchWithFallback(
    context: Word.RequestContext,
    container: WordSearchContainer,
    searchText: string,
  ): Promise<Word.Range | null> {
    return this.textLocator.locate({ context, container, searchText });
  }

  /**
   * Re-locates a suggestion by contextual anchor when direct artifact lookup is
   * missing.
   *
   * This method deliberately delegates to `SuggestionTextRangeLocator` instead
   * of searching the anchor directly in the body. A global anchor search can
   * select a table-of-contents or heading occurrence that merely shares the same
   * word, which is worse than not navigating.
   */
  async resolveSuggestionFallbackRange(
    context: Word.RequestContext,
    suggestion: Suggestion,
  ): Promise<Word.Range | null> {
    return this.textRangeLocator.locateAnchorInContext(
      context,
      context.document.body,
      suggestion,
      { logPrefix: "🧭 [WordSuggestionNavigationAdapter]" },
    );
  }

  /** Selects a range if present so Word scrolls the viewport to it. */
  async selectRange(
    context: Word.RequestContext,
    range: Word.Range | null,
  ): Promise<void> {
    if (!range) {
      return;
    }

    range.select();
    await context.sync();
  }

  /**
   * Navigates to the real suggestion artifact or strict contextual fallback.
   *
   * The method never throws to the taskpane. Instead it returns a semantic
   * result so the UI can tell the user when navigation was unsafe, ambiguous, or
   * blocked by Word.
   */
  async navigateToText(
    target: Suggestion | string,
  ): Promise<SuggestionNavigationResult> {
    try {
      return await Word.run(async (context) => {
        if (typeof target !== "string") {
          return this.navigateToSuggestion(context, target);
        }

        const range = await this.searchWithFallback(
          context,
          context.document.body as unknown as WordSearchContainer,
          target,
        );

        if (!range) {
          return { status: "not-found", reason: "plain-text-not-found" };
        }

        await this.selectRange(context, range);
        return { status: "navigated" };
      });
    } catch {
      return { status: "failed", reason: "word-error" };
    }
  }

  /**
   * Navigates to one suggestion using artifact identity before textual fallback.
   *
   * Track-change suggestions use the operational wrapper identity. Comment-only
   * suggestions use the canonical comment-only Content Control. Textual fallback
   * runs only when no artifact exists; malformed or ambiguous artifacts fail
   * closed and do not fall through to fuzzy text matching.
   */
  private async navigateToSuggestion(
    context: Word.RequestContext,
    suggestion: Suggestion,
  ): Promise<SuggestionNavigationResult> {
    const artifact =
      suggestion.type === "comment-only"
        ? await this.artifactLocator.locateCommentOnlyArtifact(
            context,
            suggestion,
          )
        : await this.artifactLocator.locateOperationalWrapper(
            context,
            suggestion,
          );

    if (artifact.selectedCc) {
      await this.selectRange(context, artifact.selectedCc.getRange());
      return { status: "navigated" };
    }

    if (artifact.locateStatus !== "cc-not-found") {
      return this.toAmbiguousNavigationResult(artifact.locateStatus);
    }

    const fallbackRange = await this.resolveSuggestionFallbackRange(
      context,
      suggestion,
    );
    if (!fallbackRange) {
      return { status: "not-found", reason: "context-not-found" };
    }

    await this.selectRange(context, fallbackRange);
    return { status: "navigated" };
  }

  /** Converts strict artifact lookup failures into navigation-safe no-op results. */
  private toAmbiguousNavigationResult(
    locateStatus: SuggestionObservationStatus | "cc-not-found",
  ): SuggestionNavigationResult {
    if (locateStatus === "identity-lost") {
      return { status: "ambiguous", reason: "identity-lost" };
    }

    if (locateStatus === "mixed-group") {
      return { status: "ambiguous", reason: "mixed-group" };
    }

    return { status: "ambiguous", reason: "multiple-artifacts" };
  }
}
