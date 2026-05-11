import type { Suggestion } from "../../domain/suggestion/Suggestion.types";
import { TrackChangeSubtypeResolver } from "./apply-suggestion/TrackChangeSubtypeResolver";

const subtypeResolver = new TrackChangeSubtypeResolver();

/** Resolves the normalized persisted track-change subtype for one suggestion. */
export function resolvePersistedTrackChangeSubtype(suggestion: Suggestion): string {
  const subtypeResolution = subtypeResolver.resolve(suggestion);
  return subtypeResolution.subtype === "insert" ? "replace" : subtypeResolution.subtype;
}
