import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import { SuggestionTextRangeLocator } from "../location/SuggestionTextRangeLocator";
import type { TextLocator } from "../WordTextLocatorContext.types";

/**
 * Resolves the live Word range for the suggestion anchor.
 */
export class ApplySuggestionAnchorResolver {
  private readonly textRangeLocator: SuggestionTextRangeLocator;

  constructor(
    private readonly suggestion: Suggestion,
    textLocator: TextLocator,
    private readonly commandId: string
  ) {
    this.textRangeLocator = new SuggestionTextRangeLocator(textLocator);
  }

  /**
   * Resolves the exact anchor range by first locating the surrounding context,
   * then searching the anchor within that context range.
   */
  async resolveAnchorRange(
    context: Word.RequestContext,
    body: Word.Body
  ): Promise<Word.Range | null> {
    return this.textRangeLocator.locateAnchorInContext(context, body, this.suggestion, {
      commandId: this.commandId,
      logPrefix: "🔬 [ApplySuggestionCommand]",
    });
  }
}
