import type { SuggestionState } from "./Suggestion.types";

/** Valid transitions for one taskpane suggestion-card lifecycle. */
export const SUGGESTION_STATE_TRANSITIONS: Record<SuggestionState, SuggestionState[]> = {
  pending: ["resolving"],
  resolving: [
    "accepted",
    "rejected",
    "unobservable",
    "identity-lost",
    "ambiguous-location",
    "mixed-group",
    "error",
  ],
  accepted: [],
  rejected: [],
  unobservable: ["resolving"],
  "identity-lost": [],
  "ambiguous-location": [],
  "mixed-group": [],
  error: ["resolving"],
};
