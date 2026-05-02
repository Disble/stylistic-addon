import type {
  Suggestion,
  TrackChangeSuggestionSubtype,
} from "../../../domain/suggestion/Suggestion.types";

/** Formatting operation encoded by backend markdown in `suggestedText`. */
export type TrackChangeFormattingInstruction = {
  kind: "italic" | "bold";
  text: string;
};

/** Full adapter-level classification for a track-change suggestion. */
export type TrackChangeSubtypeResolution =
  | {
      subtype: TrackChangeSuggestionSubtype;
      formatting?: TrackChangeFormattingInstruction;
    }
  | { subtype: "insert" };

/**
 * Resolves the native Word Track Changes subtype represented by a suggestion.
 *
 * The backend currently encodes typography as markdown in `suggestedText`; this
 * resolver turns that transport encoding into an explicit adapter subtype so the
 * apply command never inserts literal `*...*` / `**...**` text into Word.
 */
export class TrackChangeSubtypeResolver {
  /** Classifies a suggestion without touching Word host state. */
  resolve(suggestion: Suggestion): TrackChangeSubtypeResolution {
    const hasAnchor = suggestion.anchor.length > 0;
    const suggestedText = suggestion.suggestedText ?? "";

    if (!hasAnchor && suggestedText.length > 0) {
      return { subtype: "insert" };
    }

    if (hasAnchor && suggestedText.length === 0) {
      return { subtype: "delete-only" };
    }

    const formatting = this.parseFormattingInstruction(suggestion);
    if (formatting) {
      return { subtype: "formatting", formatting };
    }

    return { subtype: "replace" };
  }

  /** Extracts a supported markdown formatting instruction for the exact anchor. */
  parseFormattingInstruction(
    suggestion: Suggestion,
  ): TrackChangeFormattingInstruction | null {
    const suggestedText = suggestion.suggestedText ?? "";

    if (
      suggestedText.startsWith("**") &&
      suggestedText.endsWith("**") &&
      suggestedText.length > 4
    ) {
      const text = suggestedText.slice(2, -2);
      return text === suggestion.anchor ? { kind: "bold", text } : null;
    }

    if (
      suggestedText.startsWith("*") &&
      suggestedText.endsWith("*") &&
      !suggestedText.startsWith("**") &&
      !suggestedText.endsWith("**") &&
      suggestedText.length > 2
    ) {
      const text = suggestedText.slice(1, -1);
      return text === suggestion.anchor ? { kind: "italic", text } : null;
    }

    return null;
  }
}
