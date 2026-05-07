import { create } from "zustand";
import type {
  ApplySuggestionsResult,
  SuggestionApplicationFailure,
} from "../domain/DocumentApplication.types";
import { mapResultStatusToState } from "../domain/suggestion/SuggestionStateMachine";
import type {
  Suggestion,
  SuggestionNavigationResult,
  SuggestionState,
} from "../domain/suggestion/Suggestion.types";
import type { SuggestionResolutionMediatorResult } from "../domain/suggestion/SuggestionResolutionWorkflow.types";
import {
  applySuggestionProgressOutcome,
  buildSuggestionProgressSummaryText,
  createSuggestionProgressSummaryModel,
} from "./SuggestionProgressSummary";
import type { ResultsPanelDeps } from "./SuggestionCardRenderer.types";
import {
  setTaskpaneCleanupVisible,
  setTaskpaneDisableTrackChangesCtaVisible,
  showTaskpaneStatus,
} from "./TaskpaneShellStore";
import { INITIAL_RESULTS_PANEL_STATE } from "./ResultsPanelStore.constants";
import type { ResultsPanelFilter } from "./ResultsPanelFilters.types";
import type {
  ResultsCardGroup,
  ResultsPanelCardState,
  ResultsPanelContext,
  ResultsPanelState,
} from "./ResultsPanelStore.types";

let context: ResultsPanelContext = { isSelection: false };

/** Zustand store holding the reactive results-panel state. */
export const useResultsPanelStore = create<ResultsPanelState>()(() => INITIAL_RESULTS_PANEL_STATE);

/** Returns the current public results-panel snapshot. */
export function getResultsPanelState(): ResultsPanelState {
  return useResultsPanelStore.getState();
}

/** Initializes the results-panel store from one pipeline completion payload. */
export function setResultsPanelData(
  suggestions: Suggestion[],
  result: ApplySuggestionsResult,
  chunkErrors: string[],
  isSelection: boolean,
  deps: ResultsPanelDeps
): void {
  const summaryModel = createSuggestionProgressSummaryModel(suggestions, result, chunkErrors);
  const cards = createInitialCards(suggestions, result.failedSuggestions);

  context = { deps, isSelection, summaryModel };
  useResultsPanelStore.setState(
    {
      activeFilter: "all",
      cards: sortCards(cards),
      summaryText: buildSuggestionProgressSummaryText(summaryModel, isSelection),
      visible: true,
    },
    true
  );

  setTaskpaneDisableTrackChangesCtaVisible(
    result.documentState === "ready-to-disable-track-changes"
  );
}

/** Clears the panel back to its initial hidden state. */
export function resetResultsPanelState(): void {
  context = { isSelection: false };
  useResultsPanelStore.setState(INITIAL_RESULTS_PANEL_STATE, true);
}

/** Hides the panel while preserving the last rendered card snapshot. */
export function hideResultsPanel(): void {
  useResultsPanelStore.setState({ visible: false });
}

/** Selects which chip filter the user wants applied to the visible cards. */
export function setResultsPanelFilter(filter: ResultsPanelFilter): void {
  useResultsPanelStore.setState({ activeFilter: filter });
}

/** Toggles the feedback accordion for one card. */
export function toggleResultsPanelFeedback(cardId: string): void {
  updateResultsPanelCards((cards) =>
    cards.map((card) =>
      card.suggestion.id === cardId ? { ...card, feedbackOpen: !card.feedbackOpen } : card
    )
  );
}

/** Updates the free-text feedback comment for one card. */
export function setResultsPanelFeedbackComment(cardId: string, feedbackComment: string): void {
  updateResultsPanelCards((cards) =>
    cards.map((card) => (card.suggestion.id === cardId ? { ...card, feedbackComment } : card))
  );
}

/** Handles safe navigation feedback for one suggestion card. */
export async function navigateResultsPanelSuggestion(cardId: string): Promise<void> {
  const deps = context.deps;
  const card = findResultsPanelCard(cardId);
  if (!deps || !card || card.isFailed) {
    return;
  }

  const navigationResult = await deps.navigateToText(card.suggestion);
  const navigationNote = getNavigationNote(navigationResult);

  updateResultsPanelCards((cards) =>
    cards.map((entry) => (entry.suggestion.id === cardId ? { ...entry, navigationNote } : entry))
  );
}

/** Handles an accept action while preserving legacy workflow semantics. */
export async function acceptResultsPanelSuggestion(cardId: string): Promise<void> {
  await resolveResultsPanelSuggestion(cardId, "accept");
}

/** Handles a reject action while preserving legacy workflow semantics. */
export async function rejectResultsPanelSuggestion(cardId: string): Promise<void> {
  await resolveResultsPanelSuggestion(cardId, "reject");
}

function createInitialCards(
  suggestions: Suggestion[],
  failedSuggestions: SuggestionApplicationFailure[]
): ResultsPanelCardState[] {
  return suggestions.map((suggestion) => {
    const failure = failedSuggestions.find((entry) => entry.suggestion.id === suggestion.id);
    const isNotFoundFailure = failure?.reason === "not-found";

    return {
      cardGroup: isNotFoundFailure ? "not-found" : "active",
      feedbackComment: "",
      feedbackOpen: false,
      failure,
      hideActions: Boolean(failure),
      isFailed: Boolean(failure),
      isNotFoundFailure,
      isResolving: false,
      state: "pending",
      suggestion,
    };
  });
}

function findResultsPanelCard(cardId: string): ResultsPanelCardState | undefined {
  return useResultsPanelStore.getState().cards.find((card) => card.suggestion.id === cardId);
}

function getNavigationNote(result: SuggestionNavigationResult): string | undefined {
  if (result.status === "navigated") {
    return undefined;
  }

  return "(no se pudo ubicar la sugerencia de forma segura)";
}

function getResolutionNote(
  nextState: SuggestionState,
  result: SuggestionResolutionMediatorResult
): string | undefined {
  switch (nextState) {
    case "identity-lost":
      return "(metadata inconsistente; reanalizá la sugerencia)";
    case "ambiguous-location":
    case "mixed-group":
      return "(resolución ambigua; reanalizá la sugerencia)";
    case "error":
      if (result.status === "cc-not-found") {
        return "(aplicación falló)";
      }
      return undefined;
    default:
      return undefined;
  }
}

function sortCards(cards: readonly ResultsPanelCardState[]): ResultsPanelCardState[] {
  const groups: Record<ResultsCardGroup, ResultsPanelCardState[]> = {
    active: [],
    processed: [],
    "not-found": [],
  };

  for (const card of cards) {
    groups[card.cardGroup].push(card);
  }

  return [...groups.active, ...groups.processed, ...groups["not-found"]];
}

async function resolveResultsPanelSuggestion(
  cardId: string,
  action: "accept" | "reject"
): Promise<void> {
  const deps = context.deps;
  const card = findResultsPanelCard(cardId);
  if (!deps || !card || card.isFailed || card.isResolving || card.hideActions) {
    return;
  }

  updateResultsPanelCards((cards) =>
    cards.map((entry) =>
      entry.suggestion.id === cardId
        ? { ...entry, isResolving: true, navigationNote: undefined }
        : entry
    )
  );

  let result: SuggestionResolutionMediatorResult;

  try {
    result =
      action === "accept"
        ? await deps.acceptSuggestion(card.suggestion, getFeedbackComment(cardId))
        : await deps.rejectSuggestion(card.suggestion, getFeedbackComment(cardId));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);

    showTaskpaneStatus(message || "Error desconocido al resolver sugerencia", "error");
    updateResultsPanelCards((cards) =>
      cards.map((entry) =>
        entry.suggestion.id === cardId
          ? {
              ...entry,
              cardGroup: "active",
              hideActions: false,
              isResolving: false,
              resolutionNote: undefined,
              state: "error",
            }
          : entry
      )
    );
    return;
  }

  const nextState = mapResultStatusToState(result.status);

  if (context.summaryModel) {
    applySuggestionProgressOutcome(context.summaryModel, cardId, result.status);
  }

  setTaskpaneDisableTrackChangesCtaVisible(result.taskpaneState.showDisableTrackChangesCta);
  setTaskpaneCleanupVisible(result.taskpaneState.showCleanupSection);

  if (
    nextState === "identity-lost" ||
    nextState === "ambiguous-location" ||
    nextState === "mixed-group"
  ) {
    showTaskpaneStatus(
      result.error ?? "La identidad persistida de la sugerencia no permite resolver con seguridad.",
      "error"
    );
  }

  if (nextState === "unobservable") {
    showTaskpaneStatus(
      result.error ?? "No se pudo confirmar el estado de la sugerencia en Word. Reintentá.",
      "error"
    );
  }

  if (nextState === "error" && result.status !== "cc-not-found") {
    showTaskpaneStatus(result.error ?? "Error desconocido al resolver sugerencia", "error");
  }

  updateResultsPanelCards((cards) =>
    cards.map((entry) => {
      if (entry.suggestion.id !== cardId) {
        return entry;
      }

      const isTerminalProcessed =
        nextState === "accepted" ||
        nextState === "rejected" ||
        nextState === "identity-lost" ||
        nextState === "ambiguous-location" ||
        nextState === "mixed-group" ||
        result.status === "cc-not-found";

      return {
        ...entry,
        cardGroup: isTerminalProcessed ? "processed" : entry.cardGroup,
        hideActions: isTerminalProcessed,
        isResolving: false,
        resolutionNote: getResolutionNote(nextState, result),
        state: nextState,
      };
    })
  );
}

function getFeedbackComment(cardId: string): string | undefined {
  const feedbackComment = findResultsPanelCard(cardId)?.feedbackComment.trim();
  return feedbackComment ? feedbackComment : undefined;
}

function updateResultsPanelCards(
  updater: (cards: readonly ResultsPanelCardState[]) => ResultsPanelCardState[]
): void {
  useResultsPanelStore.setState((state) => {
    const cards = updater(state.cards);
    const summaryText = context.summaryModel
      ? buildSuggestionProgressSummaryText(context.summaryModel, context.isSelection)
      : state.summaryText;

    return {
      cards: sortCards(cards),
      summaryText,
    };
  });
}
