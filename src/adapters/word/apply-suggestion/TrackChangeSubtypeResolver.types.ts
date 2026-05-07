import type { TrackChangeSuggestionSubtype } from "../../../domain/suggestion/Suggestion.types";

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
