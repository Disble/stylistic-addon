/**
 * Task-pane — Composition Root and top-level event handlers.
 *
 * Responsibilities:
 * - Instantiate all adapters, decorators, mediators, and the pipeline.
 * - Initialize after the React shell is mounted and Office confirms the host is Word.
 * - Expose top-level workflow handlers to the React shell.
 * - Register a `PipelineObserver` to relay pipeline events into presentation stores.
 *
 * React now owns shell/results rendering and UI event wiring. This module still owns
 * orchestration and publication into presentation stores/facades.
 *
 * @module taskpane
 */

import { HubAnalysisAdapter } from "../adapters/hub/HubAnalysisAdapter";
import { HubFeedbackAdapter } from "../adapters/hub/HubFeedbackAdapter";
import { HubUserCorrectionPreferencesAdapter } from "../adapters/hub/HubUserCorrectionPreferencesAdapter";
import { ConsoleResolutionObservabilityAdapter } from "../adapters/observability/ConsoleResolutionObservabilityAdapter";
import { OfficeUserPreferencesAdapter } from "../adapters/preferences/OfficeUserPreferencesAdapter";
import { RetryAnalysisDecorator } from "../adapters/RetryAnalysisDecorator";
import { WordAdapter } from "../adapters/word/WordAdapter";
import { AnalyzeChunksHandler } from "../domain/pipeline/handlers/AnalyzeChunksHandler";
import { ApplySuggestionsHandler } from "../domain/pipeline/handlers/ApplySuggestionsHandler";
import { CheckConnectionHandler } from "../domain/pipeline/handlers/CheckConnectionHandler";
import { ChunkTextHandler } from "../domain/pipeline/handlers/ChunkTextHandler";
import { DeduplicateHandler } from "../domain/pipeline/handlers/DeduplicateHandler";
import { GuardAppliedHandler } from "../domain/pipeline/handlers/GuardAppliedHandler";
import { ReadTextHandler } from "../domain/pipeline/handlers/ReadTextHandler";
import { PipelineOrchestrator } from "../domain/pipeline/PipelineOrchestrator";
import { PipelineStateMachine } from "../domain/pipeline/PipelineStateMachine";
import type { AnalysisProfileId } from "../domain/Profile.types";
import type {
  IFeedbackPort,
  IUserCorrectionPreferencesPort,
  IUserPreferencesPort,
} from "../domain/ports";
import type { UserCorrectionPreferences } from "../domain/user-preferences/UserCorrectionPreferences.types";
import { ReviewSessionMediator } from "../domain/review/ReviewSessionMediator";
import { HttpClient } from "../infrastructure/http/HttpClient";
import { buildHubApiBaseUrl, readHubToken } from "../infrastructure/hub/HubConnection.helpers";
import { HUB_TOKEN_MISSING_MESSAGE } from "./taskpane.constants";
import { DEFAULT_PROFILES, MAX_RETRIES, RETRY_BASE_DELAY_MS } from "../infrastructure/config";
import { setSelectionPreviewSnapshot } from "./SelectionPreviewStore";
import type { ResultsPanelDeps } from "./SuggestionCardRenderer.types";
import {
  getTaskpaneShellState,
  setTaskpaneSelectedGenero,
  showTaskpaneStatus,
} from "./TaskpaneShellStore";
import {
  handleAnalyze as runTaskpaneAnalysis,
  handleCancelAnalysis as cancelTaskpaneAnalysis,
  handleRetryAnalysisQuery as retryTaskpaneAnalysisQuery,
} from "./TaskpaneAnalysisHandlers";
import type { TaskpaneAnalysisHandlersRuntime } from "./TaskpaneAnalysisHandlers.types";
import {
  bootstrapTaskpane as bootstrapTaskpaneRuntime,
  handleCleanup as cleanupTaskpaneRuntime,
  handleDisableTrackChanges as disableTrackChangesTaskpaneRuntime,
  handleLoadPreferences as loadTaskpanePreferences,
  handleSavePreferences as saveTaskpanePreferences,
  refreshCleanupVisibility,
} from "./TaskpaneBootstrap";
import type { TaskpaneBootstrapRuntime } from "./TaskpaneBootstrap.types";

// ---------------------------------------------------------------------------
// Infrastructure — built once, reused across pipeline runs
// ---------------------------------------------------------------------------

const observabilityPort = new ConsoleResolutionObservabilityAdapter();
const documentPort = new WordAdapter(undefined, observabilityPort);
const userPreferencesPort: IUserPreferencesPort = new OfficeUserPreferencesAdapter();
const hubToken = readHubToken(globalThis.document);
const hubHttpClient = new HttpClient({
  baseUrl: buildHubApiBaseUrl(globalThis.location?.origin ?? ""),
  getAuthToken: () => hubToken,
});
const userCorrectionPreferencesPort: IUserCorrectionPreferencesPort =
  new HubUserCorrectionPreferencesAdapter(hubHttpClient);
const analysisPort = new RetryAnalysisDecorator(
  new HubAnalysisAdapter(hubHttpClient),
  MAX_RETRIES,
  RETRY_BASE_DELAY_MS
);

const feedbackPort: IFeedbackPort = new HubFeedbackAdapter(hubHttpClient);
const reviewSessionMediator = new ReviewSessionMediator(documentPort, feedbackPort);

const orchestrator = new PipelineOrchestrator([
  new ReadTextHandler(),
  new CheckConnectionHandler(),
  new ChunkTextHandler(),
  new AnalyzeChunksHandler(),
  new DeduplicateHandler(),
  new GuardAppliedHandler(),
  new ApplySuggestionsHandler(),
]);

const stateMachine = new PipelineStateMachine();
const cancelState = { value: false };

/** Deps injected into the card renderer — closures over module-level ports. */
const cardRendererDeps: ResultsPanelDeps = {
  navigateToText: (target) => documentPort.navigateToText(target),
  acceptSuggestion: (s, comment) => reviewSessionMediator.acceptSuggestion(s, comment),
  rejectSuggestion: (s, comment) => reviewSessionMediator.rejectSuggestion(s, comment),
};

/** Shared runtime injected into extracted taskpane analysis handlers. */
const analysisHandlersRuntime: TaskpaneAnalysisHandlersRuntime = {
  analysisPort,
  cardRendererDeps,
  cancelState,
  documentPort,
  getSelectedGenero: () => getTaskpaneShellState().selectedGenero,
  orchestrator,
  refreshCleanupVisibility: () => refreshCleanupVisibility(bootstrapRuntime),
  stateMachine,
};

/** Shared runtime injected into bootstrap and settings handlers. */
const bootstrapRuntime: TaskpaneBootstrapRuntime = {
  documentPort,
  onSelectionSnapshot: setSelectionPreviewSnapshot,
  reviewSessionMediator,
  setSelectedGenero: setTaskpaneSelectedGenero,
  supportedAnalysisProfiles: DEFAULT_PROFILES.map((profile) => profile.id),
  userCorrectionPreferencesPort,
  userPreferencesPort,
};

// ---------------------------------------------------------------------------
// Initialization
// ---------------------------------------------------------------------------

/**
 * Initializes taskpane DOM bindings once the React shell already exists.
 * Rehydrates shell visibility from document state after the React shell mounts.
 */
export function bootstrapTaskpane(): void {
  bootstrapTaskpaneRuntime(bootstrapRuntime);
}

/**
 * Loads the user's correction-instruction preferences from the backend.
 * Used by the settings page to hydrate its draft state when it opens.
 */
export async function handleLoadPreferences(): Promise<UserCorrectionPreferences> {
  return loadTaskpanePreferences(bootstrapRuntime);
}

/**
 * Persists the full settings form atomically:
 * - PUT correction instructions to the backend (most likely to fail).
 * - Persist the analysis profile in OfficeRuntime storage.
 * - Commit the new profile into the shell store on success.
 *
 * Throws on any failure so the settings page can render an inline message.
 */
export async function handleSavePreferences(
  correctionInstructions: string | null,
  analysisProfile: AnalysisProfileId
): Promise<UserCorrectionPreferences> {
  return saveTaskpanePreferences(bootstrapRuntime)(correctionInstructions, analysisProfile);
}

// ---------------------------------------------------------------------------
// Main Event Handler — Analysis Pipeline
// ---------------------------------------------------------------------------

/**
 * Handles the "Analizar y sugerir" button click.
 *
 * Creates the `PipelineContext` with injected adapters and a UI observer,
 * then delegates entirely to `PipelineOrchestrator.run()`.
 * All UI updates come back via `PipelineEventEmitter` (Observer pattern).
 */
export async function handleAnalyze(): Promise<void> {
  if (!hubToken) {
    showTaskpaneStatus(HUB_TOKEN_MISSING_MESSAGE, "error");
    return;
  }

  return runTaskpaneAnalysis(analysisHandlersRuntime);
}

// ---------------------------------------------------------------------------
// Comment Cleanup Handler
// ---------------------------------------------------------------------------

/** Cleans up resolved Stylistic comments from the active document. */
export async function handleCleanup(): Promise<void> {
  return cleanupTaskpaneRuntime(bootstrapRuntime);
}

// ---------------------------------------------------------------------------
// Disable Track Changes Handler
// ---------------------------------------------------------------------------

/** Disables Track Changes through the review-session workflow. */
export async function handleDisableTrackChanges(): Promise<void> {
  return disableTrackChangesTaskpaneRuntime(bootstrapRuntime);
}

/** Cancels every currently active backend run in the in-memory taskpane session. */
export async function handleCancelAnalysis(): Promise<void> {
  return cancelTaskpaneAnalysis(analysisHandlersRuntime);
}

/** Re-runs backend polling only for already submitted run ids. */
export async function handleRetryAnalysisQuery(): Promise<void> {
  return retryTaskpaneAnalysisQuery(analysisHandlersRuntime);
}
