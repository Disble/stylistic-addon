import type { SuggestionObservationStatus } from "../../../domain/suggestion/Suggestion.types";

/** Result of locating persisted Word artifacts for one suggestion. */
export type LocatedSuggestionArtifactResult = {
  /** All Content Controls returned by Word for the exact lookup tag. */
  candidates: Word.ContentControl[];

  /** The single safe Content Control candidate, or null when lookup is unsafe. */
  selectedCc: Word.ContentControl | null;

  /** Semantic lookup status consumed by navigation and resolution workflows. */
  locateStatus: SuggestionObservationStatus | "cc-not-found";
};
