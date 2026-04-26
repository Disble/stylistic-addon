/* global Word */

import type { Suggestion } from "../../domain/suggestion/Suggestion.types";
import { STYLISTIC_TAG_PREFIX } from "../../infrastructure/config";
import {
  isValidOperationalReplaceIdentity,
  parseReplaceIdentityTitle,
} from "./ReplaceIdentityParser";
import type {
  TextLocator,
  WordSearchContainer,
} from "./WordTextLocatorContext";

/** Handles artifact-first navigation and resilient fallback search in Word. */
export class WordSuggestionNavigationAdapter {
  constructor(private readonly textLocator: TextLocator) {}

  /** Returns the canonical Stylistic tag for one suggestion. */
  buildSuggestionTag(suggestion: Suggestion): string {
    return `${STYLISTIC_TAG_PREFIX}${suggestion.type}:${suggestion.id}`;
  }

  /** Chooses a strict operational-wrapper navigation CC without ranking fallback. */
  selectNavigationContentControl(
    ccs: Word.ContentControl[],
    suggestion: Suggestion,
  ): Word.ContentControl | null {
    if (ccs.length === 0) {
      return null;
    }

    const validCandidates = ccs.filter((cc) => {
      const identity = parseReplaceIdentityTitle(cc.title);
      return isValidOperationalReplaceIdentity(identity, suggestion);
    });

    return validCandidates.length === 1 ? validCandidates[0] : null;
  }

  /** Searches a body or range through the shared Word text locator. */
  async searchWithFallback(
    context: Word.RequestContext,
    container: WordSearchContainer,
    searchText: string,
  ): Promise<Word.Range | null> {
    return this.textLocator.locate({ context, container, searchText });
  }

  /** Re-locates a suggestion by contextual anchor when direct artifact lookup fails. */
  async resolveSuggestionFallbackRange(
    context: Word.RequestContext,
    suggestion: Suggestion,
  ): Promise<Word.Range | null> {
    const body = context.document.body as unknown as WordSearchContainer;
    const contextRange = await this.searchWithFallback(
      context,
      body,
      suggestion.context,
    );

    if (!contextRange) {
      return this.searchWithFallback(context, body, suggestion.anchor);
    }

    contextRange.load("text");
    const containingParagraph = contextRange.paragraphs
      .getFirst()
      .getRange("Whole");
    containingParagraph.load("text");
    await context.sync();

    const shouldExpandToParagraph =
      !contextRange.text.includes(suggestion.anchor) &&
      contextRange.text.length < suggestion.context.length - 20;

    const searchContainer = shouldExpandToParagraph
      ? (containingParagraph as unknown as WordSearchContainer)
      : (contextRange as unknown as WordSearchContainer);

    return this.searchWithFallback(context, searchContainer, suggestion.anchor);
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

  /** Navigates to the real suggestion artifact or falls back to text search. */
  async navigateToText(target: Suggestion | string): Promise<void> {
    try {
      await Word.run(async (context) => {
        if (typeof target !== "string") {
          const ccResult = context.document.contentControls.getByTag(
            this.buildSuggestionTag(target),
          );
          ccResult.load("items/tag,items/title");
          await context.sync();

          const selectedCc = this.selectNavigationContentControl(
            ccResult.items,
            target,
          );
          if (selectedCc) {
            await this.selectRange(context, selectedCc.getRange());
            return;
          }

          await this.selectRange(
            context,
            await this.resolveSuggestionFallbackRange(context, target),
          );
          return;
        }

        await this.selectRange(
          context,
          await this.searchWithFallback(
            context,
            context.document.body as unknown as WordSearchContainer,
            target,
          ),
        );
      });
    } catch {
      // Navigation is best-effort — silently ignore all failures
    }
  }
}
