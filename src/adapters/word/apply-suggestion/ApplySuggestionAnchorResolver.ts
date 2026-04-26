/* global console */

import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import type {
  TextLocator,
  WordSearchContainer,
} from "../WordTextLocatorContext";

/**
 * Resolves the live Word range for the suggestion anchor.
 */
export class ApplySuggestionAnchorResolver {
  constructor(
    private readonly suggestion: Suggestion,
    private readonly textLocator: TextLocator,
    private readonly commandId: string,
  ) {}

  /**
   * Resolves the exact anchor range by first locating the surrounding context,
   * then searching the anchor within that context range.
   */
  async resolveAnchorRange(
    context: Word.RequestContext,
    body: Word.Body,
  ): Promise<Word.Range | null> {
    const contextRange = await this.textLocator.locate({
      context,
      container: body as WordSearchContainer,
      searchText: this.suggestion.context,
    });
    if (!contextRange) {
      console.log(
        `🔬 [ApplySuggestionCommand] "${this.commandId}": context not found — ambiguous-location abort before mutation`,
      );
      return null;
    }

    contextRange.load("text");
    const containingParagraph = contextRange.paragraphs
      .getFirst()
      .getRange("Whole");
    containingParagraph.load("text");
    await context.sync();

    const matchText = contextRange.text;
    console.log(
      `🔬 [ApplySuggestionCommand] "${this.commandId}": contextMatchLen=${matchText.length}, paragraphLen=${containingParagraph.text.length}, anchorIndexInMatch=${matchText.indexOf(this.suggestion.anchor)}, anchorIndexInParagraph=${containingParagraph.text.indexOf(this.suggestion.anchor)}`,
    );

    const shouldExpandToParagraph =
      !matchText.includes(this.suggestion.anchor) &&
      matchText.length < this.suggestion.context.length - 20;
    const shouldRetryInParagraphAfterMiss =
      !shouldExpandToParagraph &&
      matchText.length < this.suggestion.context.length - 20 &&
      containingParagraph.text.length > matchText.length;

    const searchContainer = shouldExpandToParagraph
      ? (containingParagraph as unknown as WordSearchContainer)
      : (contextRange as unknown as WordSearchContainer);

    if (shouldExpandToParagraph) {
      console.log(
        `🔬 [ApplySuggestionCommand] "${this.commandId}": context match (${matchText.length} chars) does not contain anchor — expanding to paragraph (${containingParagraph.text.length} chars)`,
      );
    }

    const anchorRange = await this.textLocator.locate({
      context,
      container: searchContainer,
      searchText: this.suggestion.anchor,
    });

    if (anchorRange || !shouldRetryInParagraphAfterMiss) {
      return anchorRange;
    }

    console.log(
      `🔬 [ApplySuggestionCommand] "${this.commandId}": anchor not found inside partial context match (${matchText.length} chars) — retrying in paragraph (${containingParagraph.text.length} chars)`,
    );

    return this.textLocator.locate({
      context,
      container: containingParagraph as unknown as WordSearchContainer,
      searchText: this.suggestion.anchor,
    });
  }
}
