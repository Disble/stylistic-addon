import type {
  ApplyMutationPatch,
  ChangeType,
} from "../../../domain/DocumentApplication.types";
import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";

/**
 * Encapsulates change classification and localized patch generation.
 */
export class ApplySuggestionMutationPatchBuilder {
  /**
   * Determines the type of tracked change operation for a suggestion.
   *
   * Must only be called for `track-change` suggestions where `suggestedText`
   * belongs to the mutation branch. Comment-only suggestions are handled by the
   * command orchestrator and never reach this classifier.
   */
  classifyChange(suggestion: Suggestion): ChangeType {
    const hasOriginal = suggestion.anchor.length > 0;
    const hasSuggested = (suggestion.suggestedText?.length ?? 0) > 0;
    if (hasOriginal && !hasSuggested) return "delete";
    if (!hasOriginal && hasSuggested) return "insert";
    return "replace";
  }

  /** Builds a localized mutation patch from one successful anchor replacement. */
  buildApplyMutationPatch(
    suggestion: Suggestion,
    containerText: string,
  ): ApplyMutationPatch | undefined {
    const replacement = suggestion.suggestedText ?? "";
    const affectedStart = containerText.indexOf(suggestion.anchor);

    if (affectedStart < 0) {
      return undefined;
    }

    const affectedEnd = affectedStart + suggestion.anchor.length;

    return {
      suggestionId: suggestion.id,
      snapshotVersion: (suggestion.positionHint?.snapshotVersion ?? 0) + 1,
      paragraphId: suggestion.positionHint?.paragraphId,
      originalText: containerText,
      updatedText:
        containerText.slice(0, affectedStart) +
        replacement +
        containerText.slice(affectedEnd),
      deltaLength: replacement.length - suggestion.anchor.length,
      affectedStart,
      affectedEnd,
    };
  }

  /** Converts unknown error values into a stable, readable log message. */
  stringifyUnknownError(error: unknown): string {
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
}
