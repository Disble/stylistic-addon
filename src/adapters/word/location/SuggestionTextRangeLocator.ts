import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import type { TextLocator, WordSearchContainer } from "../WordTextLocatorContext.types";
import type { SuggestionTextRangeLocatorOptions } from "./SuggestionTextRangeLocator.types";

/**
 * Locates a suggestion anchor inside its surrounding context without mutation.
 *
 * This is the shared textual fallback for workflows that need to localize a
 * suggestion from backend text. It is strict by design: the context must be
 * found before the anchor is searched, and the anchor is searched only inside
 * the localized context or its containing paragraph. It never performs a global
 * body-wide anchor search.
 */
export class SuggestionTextRangeLocator {
  constructor(private readonly textLocator: TextLocator) {}

  /**
   * Resolves the anchor range by first locating context, then anchor within it.
   *
   * Returns `null` when the context cannot be localized or when the anchor cannot
   * be found inside the localized scope. Callers must treat `null` as an unsafe
   * location, not as permission to broaden the search globally.
   */
  async locateAnchorInContext(
    context: Word.RequestContext,
    body: Word.Body,
    suggestion: Suggestion,
    options: SuggestionTextRangeLocatorOptions = {}
  ): Promise<Word.Range | null> {
    const contextRange = await this.textLocator.locate({
      context,
      container: body as WordSearchContainer,
      searchText: suggestion.context,
    });
    if (!contextRange) {
      this.log(options, "context not found — ambiguous-location abort before action");
      return null;
    }

    contextRange.load("text");
    const containingParagraph = contextRange.paragraphs.getFirst().getRange("Whole");
    containingParagraph.load("text");
    await context.sync();

    const matchText = contextRange.text;
    const containingParagraphText = containingParagraph.text;
    this.log(
      options,
      `contextMatchLen=${matchText.length}, paragraphLen=${containingParagraphText.length}, anchorIndexInMatch=${matchText.indexOf(suggestion.anchor)}, anchorIndexInParagraph=${containingParagraphText.indexOf(suggestion.anchor)}`
    );

    const shouldExpandToParagraph =
      !matchText.includes(suggestion.anchor) && matchText.length < suggestion.context.length - 20;
    const shouldRetryInParagraphAfterMiss =
      !shouldExpandToParagraph &&
      matchText.length < suggestion.context.length - 20 &&
      containingParagraphText.length > matchText.length;

    const searchContainer = shouldExpandToParagraph
      ? (containingParagraph as unknown as WordSearchContainer)
      : (contextRange as unknown as WordSearchContainer);

    if (shouldExpandToParagraph) {
      this.log(
        options,
        `context match (${matchText.length} chars) does not contain anchor — expanding to paragraph (${containingParagraphText.length} chars)`
      );
    }

    const anchorRange = await this.textLocator.locate({
      context,
      container: searchContainer,
      searchText: suggestion.anchor,
    });

    if (anchorRange || !shouldRetryInParagraphAfterMiss) {
      return anchorRange;
    }

    this.log(
      options,
      `anchor not found inside partial context match (${matchText.length} chars) — retrying in paragraph (${containingParagraphText.length} chars)`
    );

    return this.textLocator.locate({
      context,
      container: containingParagraph as unknown as WordSearchContainer,
      searchText: suggestion.anchor,
    });
  }

  /** Emits optional diagnostics using the workflow-specific prefix. */
  private log(options: SuggestionTextRangeLocatorOptions, message: string): void {
    if (!options.logPrefix) {
      return;
    }

    const command = options.commandId ? ` "${options.commandId}"` : "";
    console.log(`${options.logPrefix}${command}: ${message}`);
  }
}
