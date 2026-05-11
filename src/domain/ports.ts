/**
 * Port interfaces (Hexagonal Architecture) — define the contracts between
 * the domain/application layer and the infrastructure adapters.
 *
 * These interfaces are the only cross-boundary contracts in the system:
 * - `IDocumentPort` — abstracts all Word document operations.
 * - `IAnalysisPort` — abstracts all Mastra backend operations.
 *
 * Benefits:
 * - The pipeline (`domain/pipeline/`) depends only on these interfaces.
 * - Adapters (`adapters/`) implement these interfaces.
 * - Tests can inject mock implementations without touching Office.js or Mastra.
 * - Swapping the backend (e.g., OpenAI direct) or document host (e.g., Google Docs)
 *   requires only a new adapter — zero changes to the pipeline or UI.
 *
 * @module ports
 */

import type { TextChunk } from "./chunking/TextChunk.types";
import type { ApplySuggestionsResult } from "./DocumentApplication.types";
import type { AuthSession, SocialSignInRequest } from "./auth/AuthSession.types";
import type {
  ChunkCancelResult,
  ChunkPollResult,
  ChunkRunReference,
  ChunkSubmitResult,
  WorkflowSubmitContext,
} from "./mastra/MastraWorkflow.types";
import type { ProgressCallback } from "./pipeline/PipelineEvents.types";
import type { AnalysisProfileId } from "./Profile.types";
import type { DocumentReviewState } from "./review/DocumentReviewStateMachine.types";
import type { SelectionSnapshot } from "./selection/SelectionSnapshot.types";
import type { Suggestion, SuggestionNavigationResult } from "./suggestion/Suggestion.types";
import type {
  FeedbackPayload,
  ResolutionObservabilityEvent,
  ResolutionObservabilitySnapshot,
  SuggestionActionResult,
} from "./suggestion/SuggestionResolutionWorkflow.types";
import type { TextSource } from "./TextSource.types";
import type { UserCorrectionPreferences } from "./user-preferences/UserCorrectionPreferences.types";

// ---------------------------------------------------------------------------
// Auth Ports
// ---------------------------------------------------------------------------

/**
 * Contract for Better Auth session operations.
 *
 * Implementations own remote auth protocol details; callers only deal with the
 * domain-level session and the provider sign-in URL used by the Office dialog.
 */
export interface IAuthPort {
  /** Starts the configured OAuth sign-in request and returns the provider URL. */
  createSocialSignInRequest(callbackUrl: string): Promise<SocialSignInRequest>;

  /** Resolves the current Better Auth session, optionally using a bearer token. */
  getSession(token?: string): Promise<AuthSession | undefined>;

  /** Revokes/signs out the current session. Implementations swallow transport details. */
  signOut(token?: string): Promise<void>;
}

/** Persistent storage boundary for auth sessions in the Office host. */
export interface IAuthSessionStoragePort {
  /** Restores the last persisted session, if any. */
  restore(): Promise<AuthSession | undefined>;

  /** Persists the latest valid session for future taskpane launches. */
  persist(session: AuthSession): Promise<void>;

  /** Removes any persisted session from host storage. */
  clear(): Promise<void>;
}

// ---------------------------------------------------------------------------
// User Preferences Port
// ---------------------------------------------------------------------------

/**
 * Persistent storage boundary for user-level preferences that survive across
 * documents and sessions (e.g., chosen analysis profile, future UI prefs).
 *
 * Adapters MUST be cross-document, per-user — preferences are configuration
 * about the user, not about the document. When a backend-backed implementation
 * arrives, the contract stays the same and only the adapter is swapped.
 */
export interface IUserPreferencesPort {
  /**
   * Returns the user's stored analysis-profile id, or `undefined` when no
   * preference is persisted or the persisted value is not a non-empty string.
   * Implementations MUST NOT validate the id against the domain whitelist —
   * the composition root performs that semantic check before hydrating state.
   */
  getAnalysisProfile(): Promise<string | undefined>;

  /** Persists the user's selected, already-validated analysis-profile id. */
  setAnalysisProfile(value: AnalysisProfileId): Promise<void>;
}

/**
 * Backend-backed boundary for the user's global correction-instruction
 * preferences. Distinct from {@link IUserPreferencesPort} because:
 *
 * - It is server-owned and authenticated, not OfficeRuntime-local.
 * - It carries a backend-echoed max-length contract used for UX validation.
 * - It exposes `null` semantics for the "no instructions" state, so passing
 *   `null` to {@link save} means "clear the persisted value".
 *
 * Adapters MUST translate transport-level failures into a
 * {@link UserCorrectionPreferencesError} with a domain-meaningful reason.
 */
export interface IUserCorrectionPreferencesPort {
  /** Loads the user's current correction-instruction preferences. */
  load(): Promise<UserCorrectionPreferences>;

  /**
   * Persists the user's correction instructions. Sending `null` clears the
   * stored value. The backend trims whitespace and may collapse to `null`.
   */
  save(correctionInstructions: string | null): Promise<UserCorrectionPreferences>;
}

// ---------------------------------------------------------------------------
// Document Port
// ---------------------------------------------------------------------------

/**
 * Contract for all document read/write operations.
 *
 * Implemented by `WordAdapter` for Microsoft Word via Office.js.
 * Could be implemented for Google Docs, LibreOffice, etc. without touching
 * any pipeline or UI code.
 */
export interface IDocumentPort {
  /**
   * Returns the stable document UUID persisted inside the active Word document.
   * Implementations must create and persist one when the document has not been
   * initialized yet.
   */
  getDocumentUuid(): Promise<string>;

  /**
   * Resolves the text to analyze: returns the current selection if non-empty,
   * otherwise falls back to the full document body.
   */
  getTextToAnalyze(): Promise<TextSource>;

  /**
   * Returns the set of original texts already applied as Stylistic tracked
   * deletions in the document. Used as a guard to prevent duplicate tracked
   * changes on re-run.
   */
  getAppliedOriginalTexts(): Promise<Set<string>>;

  /**
   * Applies an array of suggestions as native tracked changes with embedded
   * justification comments, one suggestion at a time.
   *
   * @param suggestions - Suggestions to apply.
   * @param onProgress  - Optional progress callback invoked after each suggestion.
   */
  applySuggestions(
    suggestions: Suggestion[],
    onProgress?: ProgressCallback
  ): Promise<ApplySuggestionsResult>;

  /**
   * Returns a dry-run summary of the cleanup operation.
   * Used by the task pane to decide whether the cleanup CTA should be visible.
   */
  getCleanupPreview(): Promise<{ deletable: number; kept: number }>;

  /**
   * Deletes Stylistic comments whose tracked changes have been resolved.
   * Never touches comments from other authors.
   */
  cleanupResolvedComments(): Promise<{ deleted: number; kept: number }>;

  /**
   * Accepts all Stylistic tracked changes associated with a suggestion.
   * Also deletes the associated Stylistic comment if present.
   * Returns a result object — never throws.
   *
   * Result contract:
   * - `accepted` means tracked changes and Stylistic cleanup completed atomically
   * - `unobservable` means Word could not prove the suggestion state yet
   * - `identity-lost` means operational-wrapper metadata exists but is incomplete/corrupt
   */
  acceptSuggestion(suggestion: Suggestion): Promise<SuggestionActionResult>;

  /**
   * Rejects all Stylistic tracked changes associated with a suggestion.
   * Also deletes the associated Stylistic comment if present.
   * Returns a result object — never throws.
   *
   * Result contract:
   * - `rejected` means tracked changes and Stylistic cleanup completed atomically
   * - `unobservable` means Word could not prove the suggestion state yet
   * - `identity-lost` means operational-wrapper metadata exists but is incomplete/corrupt
   */
  rejectSuggestion(suggestion: Suggestion): Promise<SuggestionActionResult>;

  /**
   * Inspects the document-derived Stylistic review state.
   * The document is the source of truth for pending artifacts.
   */
  getDocumentReviewState(): Promise<DocumentReviewState>;

  /**
   * Lets the user explicitly disable Word Track Changes after Stylistic pending
   * artifacts reach zero. Never called automatically.
   */
  disableTrackChanges(): Promise<void>;

  /**
   * Subscribes to host selection changes and emits a `SelectionSnapshot` whenever
   * the user updates the active selection. Implementations must emit one initial
   * snapshot synchronously or asynchronously after subscription so the UI can
   * reflect the current state.
   *
   * Returns a function that cancels the subscription.
   */
  subscribeSelectionChanges(listener: (snapshot: SelectionSnapshot) => void): () => void;

  /**
   * Navigates the document view to the real Word artifact for one suggestion.
   * Prefers persisted Stylistic identity and falls back only to strict
   * `context -> anchor` text localization when the artifact no longer exists.
   * It must never use global anchor search for suggestions because selecting the
   * wrong occurrence is more harmful than refusing to navigate.
   * Never throws. Returns a semantic result so the UI can inform the user when
   * no safe navigation target exists.
   */
  navigateToText(target: Suggestion | string): Promise<SuggestionNavigationResult>;
}

// ---------------------------------------------------------------------------
// Analysis Port
// ---------------------------------------------------------------------------

/**
 * Contract for all AI analysis backend operations.
 *
 * Implemented by `MastraAdapter` (wrapped with `RetryAnalysisDecorator`).
 * Could be implemented for any LLM backend or local analysis engine.
 */
export interface IAnalysisPort {
  /**
   * Checks whether the backend is reachable and the workflow exists.
   * Used as a fail-fast gate before starting analysis.
   * Never throws — returns `false` on any error.
   */
  checkConnection(): Promise<boolean>;

  /**
   * Submits a text chunk for asynchronous workflow execution.
   * Returns the `runId` required for later polling.
   */
  submitChunkAnalysis(chunk: TextChunk, input: WorkflowSubmitContext): Promise<ChunkSubmitResult>;

  /**
   * Polls an existing workflow run created by `submitChunkAnalysis()`.
   * Returns intermediate or terminal workflow state for the chunk.
   */
  pollChunkAnalysis(chunkIndex: number, runId: string): Promise<ChunkPollResult>;

  /**
   * Requests backend cancellation for one active workflow run.
   * Used exclusively by the taskpane cancel action while polling is active.
   */
  cancelChunkAnalysis(chunkIndex: number, runId: string): Promise<ChunkCancelResult>;

  /**
   * Polls a previously known workflow run again without resubmitting the chunk.
   * This supports frontend-only recovery when an earlier poll attempt failed
   * locally but the backend run might still be alive.
   */
  retryPollChunkAnalysis(reference: ChunkRunReference): Promise<ChunkPollResult>;
}

// ---------------------------------------------------------------------------
// Feedback Port
// ---------------------------------------------------------------------------

/**
 * Contract for sending user feedback about suggestions.
 *
 * Implemented by `FeedbackAdapter`.
 * Fire-and-forget — never awaited in the UI. Errors must be swallowed silently.
 */
export interface IFeedbackPort {
  /**
   * Sends a feedback payload to the backend.
   * Must execute asynchronously; errors must be swallowed silently.
   */
  sendFeedback(payload: FeedbackPayload): Promise<void>;
}

// ---------------------------------------------------------------------------
// Resolution Observability Port
// ---------------------------------------------------------------------------

/**
 * Contract for best-effort resolution observability emitted by workflow phases.
 *
 * Observability must NEVER alter semantic document outcomes. Adapters are
 * expected to swallow transport/storage failures internally or let callers
 * degrade them to warnings.
 */
export interface IResolutionObservabilityPort {
  /** Emits one structured phase event for later debugging or analysis. */
  emitEvent(event: ResolutionObservabilityEvent): Promise<void>;

  /** Captures one structured host-evidence snapshot for forensic debugging. */
  captureSnapshot(snapshot: ResolutionObservabilitySnapshot): Promise<void>;
}
