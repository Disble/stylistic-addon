import type {
  ApplySuggestionsResult,
  SuggestionApplicationFailure,
} from "../domain/DocumentApplication.types";
import type { Suggestion, SuggestionState } from "../domain/suggestion/Suggestion.types";
import type { ResultsPanelDeps } from "./SuggestionCardRenderer.types";
import type { SuggestionProgressSummaryModel } from "./SuggestionProgressSummary.types";
import type { ResultsPanelFilter } from "./ResultsPanelFilters.types";

/** High-level lifecycle group used to order result cards. */
export type ResultsCardGroup = "active" | "processed" | "not-found";

/** Presentation state for one rendered suggestion card. */
export type ResultsPanelCardState = Readonly<{
  cardGroup: ResultsCardGroup;
  feedbackComment: string;
  feedbackOpen: boolean;
  failure?: SuggestionApplicationFailure;
  hideActions: boolean;
  isFailed: boolean;
  isNotFoundFailure: boolean;
  isResolving: boolean;
  navigationNote?: string;
  resolutionNote?: string;
  state: SuggestionState;
  suggestion: Suggestion;
}>;

/** Root Zustand state for the taskpane results panel. */
export type ResultsPanelState = Readonly<{
  activeFilter: ResultsPanelFilter;
  cards: readonly ResultsPanelCardState[];
  summaryText: string;
  visible: boolean;
}>;

/** Private dependencies and summary model used by panel actions. */
export type ResultsPanelContext = {
  deps?: ResultsPanelDeps;
  isSelection: boolean;
  summaryModel?: SuggestionProgressSummaryModel;
};

/** Input payload used when the panel is initialized from pipeline completion. */
export type ResultsPanelDataInput = Readonly<{
  suggestions: Suggestion[];
  result: ApplySuggestionsResult;
  chunkErrors: string[];
  isSelection: boolean;
  deps: ResultsPanelDeps;
}>;
