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
  type SuggestionProgressSummaryModel,
} from "./SuggestionProgressSummary";
import type { ResultsPanelDeps } from "./SuggestionCardRenderer.types";
import {
  setTaskpaneCleanupVisible,
  setTaskpaneDisableTrackChangesCtaVisible,
  showTaskpaneStatus,
} from "./TaskpaneShellStore";

type ResultsCardGroup = "active" | "processed" | "not-found";

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

export type ResultsPanelState = Readonly<{
  cards: readonly ResultsPanelCardState[];
  summaryText: string;
  visible: boolean;
}>;

type ResultsPanelInternalState = {
  deps?: ResultsPanelDeps;
  isSelection: boolean;
  publicState: ResultsPanelState;
  summaryModel?: SuggestionProgressSummaryModel;
};

const INITIAL_PUBLIC_STATE: ResultsPanelState = {
  cards: [],
  summaryText: "",
  visible: false,
};

let internalState: ResultsPanelInternalState = {
  isSelection: false,
  publicState: INITIAL_PUBLIC_STATE,
};

const listeners = new Set<() => void>();

/** Subscribes React consumers to results-panel state changes. */
export function subscribeResultsPanelStore(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

/** Returns the current public results-panel snapshot. */
export function getResultsPanelState(): ResultsPanelState {
  return internalState.publicState;
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

  internalState = {
    deps,
    isSelection,
    summaryModel,
    publicState: {
      cards: sortCards(cards),
      summaryText: buildSuggestionProgressSummaryText(summaryModel, isSelection),
      visible: true,
    },
  };

  setTaskpaneDisableTrackChangesCtaVisible(
    result.documentState === "ready-to-disable-track-changes"
  );
  emitResultsPanelChange();
}

/** Clears the panel back to its initial hidden state. */
export function resetResultsPanelState(): void {
  internalState = {
    isSelection: false,
    publicState: INITIAL_PUBLIC_STATE,
  };
  emitResultsPanelChange();
}

/** Hides the panel while preserving the last rendered card snapshot. */
export function hideResultsPanel(): void {
  internalState = {
    ...internalState,
    publicState: {
      ...internalState.publicState,
      visible: false,
    },
  };
  emitResultsPanelChange();
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
  const deps = internalState.deps;
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

function emitResultsPanelChange(): void {
  for (const listener of listeners) {
    listener();
  }
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
  return internalState.publicState.cards.find((card) => card.suggestion.id === cardId);
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
  const deps = internalState.deps;
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

  if (internalState.summaryModel) {
    applySuggestionProgressOutcome(internalState.summaryModel, cardId, result.status);
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
  const cards = updater(internalState.publicState.cards);
  const summaryText = internalState.summaryModel
    ? buildSuggestionProgressSummaryText(internalState.summaryModel, internalState.isSelection)
    : internalState.publicState.summaryText;

  internalState = {
    ...internalState,
    publicState: {
      cards: sortCards(cards),
      summaryText,
      visible: internalState.publicState.visible,
    },
  };

  emitResultsPanelChange();
}
