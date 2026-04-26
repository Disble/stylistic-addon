import type { Suggestion } from "../domain/suggestion/Suggestion.types";
import type { SuggestionResolutionMediatorResult } from "../domain/suggestion/SuggestionResolutionWorkflow.types";
import type { SuggestionProgressSummaryModel } from "./SuggestionProgressSummary";

/** Business-layer capabilities needed to render and interact with cards. */
export type ResultsPanelDeps = {
  navigateToText: (target: Suggestion | string) => Promise<void>;
  acceptSuggestion: (
    suggestion: Suggestion,
    comment?: string,
  ) => Promise<SuggestionResolutionMediatorResult>;
  rejectSuggestion: (
    suggestion: Suggestion,
    comment?: string,
  ) => Promise<SuggestionResolutionMediatorResult>;
};

/** Shared UI state required while resolving one rendered suggestion card. */
export type SuggestionResolutionUiContext = {
  summaryModel: SuggestionProgressSummaryModel;
  summaryElement: HTMLElement;
  isSelection: boolean;
};

/** The pair of action buttons owned by one suggestion card. */
export type SuggestionActionButtons = {
  acceptBtn: HTMLButtonElement | null;
  rejectBtn: HTMLButtonElement | null;
};

/** Rendered card metadata used by list orchestration. */
export type RenderedSuggestionCard = {
  li: HTMLLIElement;
  isFailed: boolean;
  isNotFoundFailure: boolean;
  suggestion: Suggestion;
};
