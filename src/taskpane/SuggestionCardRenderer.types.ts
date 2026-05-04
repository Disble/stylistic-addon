import type { Suggestion, SuggestionNavigationResult } from "../domain/suggestion/Suggestion.types";
import type { SuggestionResolutionMediatorResult } from "../domain/suggestion/SuggestionResolutionWorkflow.types";

/** Business-layer capabilities needed to render and interact with cards. */
export type ResultsPanelDeps = {
  navigateToText: (target: Suggestion | string) => Promise<SuggestionNavigationResult>;
  acceptSuggestion: (
    suggestion: Suggestion,
    comment?: string
  ) => Promise<SuggestionResolutionMediatorResult>;
  rejectSuggestion: (
    suggestion: Suggestion,
    comment?: string
  ) => Promise<SuggestionResolutionMediatorResult>;
};
