/**
 * Shared TypeScript interfaces for the Stylistic add-in.
 *
 * These contracts define the communication boundaries between:
 * - The Mastra backend (workflow input/output)
 * - The Word API layer (suggestions and insertion results)
 * - The UI orchestrator (progress reporting)
 * - The pipeline (state machine, commands, events)
 *
 * No runtime code — only type declarations. No dependency on Office.js,
 * @mastra/client-js, or the DOM, keeping all consumers independently testable.
 *
 * @module types
 */

// ---------------------------------------------------------------------------
// Suggestion & Insertion
// ---------------------------------------------------------------------------

/** Supported editorial severity levels across suggestions and feedback. */
export type SuggestionSeverity = "high" | "medium" | "low";

/** Supported suggestion materialization modes in Word. */
export type SuggestionType = "track-change" | "comment-only";

/**
 * Frontend-owned position hint derived from a live apply snapshot.
 *
 * This is intentionally optional while the batch apply workflow migrates away
 * from backend-array heuristics toward real document-position ranking.
 */
export interface SuggestionBatchPositionHint {
  /** Start offset relative to the coordinate source that produced this hint. */
  start: number;

  /** End offset relative to the coordinate source that produced this hint. */
  end: number;

  /** Snapshot version whose coordinates this hint is comparable against. */
  snapshotVersion: number;

  /** Optional paragraph/local-container identity for localized recovery. */
  paragraphId?: string;

  /** Identifies the coordinate/ranking source that produced this hint. */
  source: "snapshot" | "localized-reread";

  /** Indicates the current hint must be revalidated with a localized reread. */
  requiresLocalReread?: boolean;
}

/**
 * A single editorial suggestion, either received from the Mastra workflow
 * or prepared for insertion into the Word document.
 */
export interface Suggestion {
  /** Unique identifier assigned by the frontend (e.g., "chunk0-3"). */
  id: string;

  /** Paragraph-level context used to locate the suggestion in the document. */
  context: string;

  /** Exact substring within `context` targeted by the suggestion. */
  anchor: string;

  /**
   * Replacement text that will appear as a tracked change.
   * Undefined when `type` is `"comment-only"` — no text replacement is made,
   * only a Word comment is inserted at the anchor location.
   */
  suggestedText?: string;

  /** Human-readable reason for the suggestion, shown in the results panel. */
  justification: string;

  /** Editorial category label (e.g., "Redundancia", "Muletilla"). */
  category: string;

  /** How critical the suggestion is. */
  severity: SuggestionSeverity;

  /**
   * Determines how the suggestion is applied to the document.
   *
   * - `"track-change"` — inserts OOXML tracked changes (`<w:del>` / `<w:ins>`)
   *   alongside a justification comment. Requires `suggestedText` to be defined.
   * - `"comment-only"` — inserts only a Word comment at the `anchor`
   *   location with no tracked change. `suggestedText` is undefined.
   */
  type: SuggestionType;

  /**
   * Optional live position hint used by batch apply orchestration.
   *
   * When present, this SHOULD outrank raw backend array order because it is
   * derived from the frontend's latest document snapshot.
   */
  positionHint?: SuggestionBatchPositionHint;
}

/**
 * Observation confidence for a suggestion materialized in Word.
 *
 * This is intentionally separate from business resolution status. Word may fail
 * to expose enough host evidence even when the suggestion still exists.
 */
export type SuggestionObservationStatus =
  | "confirmed-pending"
  | "confirmed-resolved"
  | "unobservable"
  | "identity-lost";

/**
 * A Word-host reference that helps re-locate one side of a review suggestion.
 *
 * `value` is intentionally opaque to the domain. It may contain a Content
 * Control tag, anchor text, tracked-change token, or another adapter-owned
 * locator string.
 */
export interface WordArtifactRef {
  /** Kind of Word artifact being referenced. */
  kind: "content-control" | "tracked-change" | "comment" | "anchor";

  /** Semantic role this artifact plays inside the suggestion identity. */
  role: "inserted-side" | "deleted-side" | "operational-anchor";

  /** Opaque adapter-owned value used to relocate the artifact in Word. */
  value: string;
}

/**
 * Versioned identity for replace suggestions.
 *
 * `compound-v2` records richer Word references without assuming the inserted-
 * side Content Control is the whole identity.
 */
export interface ReplaceSuggestionIdentity {
  /** Stable frontend/domain suggestion identifier. */
  suggestionId: string;

  /** Serialized identity version. */
  version: "compound-v2";

  /** Primary inserted-side Word reference. */
  insertedSideRef: WordArtifactRef;

  /** Optional deleted/original-side Word reference. */
  deletedSideRef?: WordArtifactRef;

  /** Optional operational anchor for fallback re-location. */
  anchorRef?: WordArtifactRef;
}

/**
 * Machine-readable failure reasons for one suggestion application attempt.
 */
export type SuggestionApplicationFailureReason =
  | "not-found"
  | "covered-by-existing-cc"
  | "command-error";

/**
 * One failed suggestion application with preserved reason and message.
 */
export interface SuggestionApplicationFailure {
  /** The suggestion that could not be applied. */
  suggestion: Suggestion;

  /** Stable failure reason used by UI and telemetry. */
  reason: SuggestionApplicationFailureReason;

  /** Human-readable adapter message captured at the failure boundary. */
  message: string;
}

/**
 * Result of attempting to insert suggestions as tracked changes in Word.
 */
export interface InsertionResult {
  /** Number of suggestions successfully applied as tracked changes. */
  successCount: number;

  /** Suggestions that failed to apply, with preserved failure semantics. */
  failedSuggestions: SuggestionApplicationFailure[];
}

/**
 * Document-derived review state for Stylistic artifacts currently materialized
 * in Word.
 */
export interface DocumentReviewState {
  /** Number of pending Stylistic artifacts still active in the document. */
  pendingStylisticArtifacts: number;

  /** Convenience boolean derived from `pendingStylisticArtifacts > 0`. */
  hasPendingStylisticArtifacts: boolean;

  /** Whether Word Track Changes is currently active for the document. */
  trackChangesActive: boolean;
}

/**
 * Taskpane-facing review state derived by the explicit review mediator.
 */
export interface ReviewTaskpaneState {
  /** Explicit document-review UI state currently exposed to the user. */
  documentState: import("./review/DocumentReviewStateMachine").DocumentReviewUiState;

  /** Whether the taskpane should expose the final Track Changes deactivation CTA. */
  showDisableTrackChangesCta: boolean;

  /** Whether the cleanup section should currently be visible in the taskpane. */
  showCleanupSection: boolean;
}

/**
 * Batch insertion result enriched with document-review semantics.
 *
 * The pipeline still consumes the insertion counters, while the workflow layer
 * can observe whether Track Changes was activated lazily and what pending state
 * remained in the document after the batch finished.
 */
export interface ApplySuggestionsResult extends InsertionResult {
  /** Document-derived review state after the batch finishes. */
  pendingAfter: DocumentReviewState;

  /** Explicit document-review UI state after the batch finishes. */
  documentState: import("./review/DocumentReviewStateMachine").DocumentReviewUiState;

  /**
   * `true` when this batch had to activate Track Changes before applying the
   * first real `track-change` suggestion.
   */
  trackChangesActivatedForBatch: boolean;
}

/**
 * Localized mutation patch captured after a successful real-Word apply.
 *
 * This is the minimal bridge toward the future incremental snapshot/rebase
 * workflow: one command reports what changed locally without forcing a full
 * document reread.
 */
export interface ApplyMutationPatch {
  /** Suggestion whose successful apply produced this patch. */
  suggestionId: string;

  /** Snapshot version after this successful mutation has been applied. */
  snapshotVersion: number;

  /** Optional paragraph/local-container identity for localized recovery. */
  paragraphId?: string;

  /** Paragraph or local container text before the mutation. */
  originalText: string;

  /** Paragraph or local container text after the mutation. */
  updatedText: string;

  /** Length delta produced by the replacement in the local text. */
  deltaLength: number;

  /** Snapshot-local start offset of the affected anchor. */
  affectedStart: number;

  /** Snapshot-local end offset of the original affected anchor. */
  affectedEnd: number;
}

// ---------------------------------------------------------------------------
// Text Chunking
// ---------------------------------------------------------------------------

/**
 * A chunk of document text prepared for sending to the Mastra workflow.
 * Created by the chunker when the document exceeds the maximum chunk size.
 */
export interface TextChunk {
  /** The text content of this chunk. */
  text: string;

  /** Zero-based index of this chunk within the full document. */
  index: number;

  /** Total number of chunks the document was split into. */
  total: number;

  /** Character offset where this chunk starts in the original document. */
  startOffset: number;
}

// ---------------------------------------------------------------------------
// Mastra Workflow Communication
// ---------------------------------------------------------------------------

/**
 * Input data sent to the Mastra stylistic workflow for each chunk.
 * Must match the workflow's `inputSchema` on the backend (`stylistic-workflow`).
 */
export interface WorkflowInput {
  /** Text to analyze. */
  text: string;

  /** Genre identifier matching the backend enum (e.g., "narrativa-literaria", "general"). */
  genero:
    | "narrativa-literaria"
    | "ensayo-academico"
    | "periodismo-cultural"
    | "general";

  /** Author slug in kebab-case used to load the author profile from the workspace. */
  autorSlug: string;
}

/**
 * Raw suggestion shape as returned by the Mastra workflow.
 * Does not include `id` — the frontend assigns IDs after receiving the response.
 */
export interface WorkflowSuggestion {
  /** Paragraph-level context fragment used to locate the suggestion. */
  context: string;

  /** Exact substring of `context` targeted by the suggestion. */
  anchor: string;

  /**
   * Replacement text. Absent when `type` is `"comment-only"`.
   */
  suggestedText?: string;

  /** Human-readable justification. */
  justification: string;

  /** Editorial category label. */
  category: string;

  /** How critical the suggestion is. */
  severity: SuggestionSeverity;

  /**
   * Suggestion kind as declared by the backend.
   * Defaults to `"track-change"` if absent (backwards compatibility).
   */
  type?: SuggestionType;
}

/**
 * Expected output from the Mastra editorial workflow on success.
 * Accessed via `result.result` after a workflow run completes.
 */
export interface WorkflowOutput {
  /** Array of editorial suggestions for the analyzed text. */
  suggestions: WorkflowSuggestion[];

  /** Optional warnings from the backend (e.g., "text too short for meaningful analysis"). */
  warnings?: string[];
}

/** Workflow statuses observed while polling a chunk analysis run. */
export type ChunkAnalysisStatus =
  | "running"
  | "success"
  | "failed"
  | "tripwire"
  | "suspended"
  | "waiting"
  | "pending"
  | "canceled"
  | "bailed"
  | "paused";

/** Result of submitting a chunk for asynchronous workflow execution. */
export interface ChunkSubmitResult {
  /** Zero-based index of the chunk this submission belongs to. */
  chunkIndex: number;

  /** Workflow run identifier used for later polling. */
  runId?: string;

  /** Error message when submission could not be confirmed. */
  error?: string;
}

/** Result of polling an asynchronous workflow run for a chunk. */
export interface ChunkPollResult {
  /** Zero-based index of the chunk this poll belongs to. */
  chunkIndex: number;

  /** Workflow run identifier being polled. */
  runId: string;

  /** Current lifecycle status of the workflow run. */
  status: ChunkAnalysisStatus;

  /** Suggestions returned when the run completed successfully. */
  suggestions: Suggestion[];

  /** Error message for terminal failed states. */
  error?: string;
}

// ---------------------------------------------------------------------------
// Progress Reporting
// ---------------------------------------------------------------------------

/**
 * Phases of the analysis pipeline, used for progress reporting in the UI.
 */
export type AnalysisPhase =
  | "reading"
  | "connecting"
  | "analyzing"
  | "applying"
  | "done";

/**
 * Callback signature for reporting progress during multi-phase analysis.
 *
 * @param phase   - Current phase of the pipeline.
 * @param current - Current step within the phase (1-based).
 * @param total   - Total steps in the phase.
 * @param message - Human-readable status message for display.
 */
export type ProgressCallback = (
  phase: AnalysisPhase,
  current: number,
  total: number,
  message: string,
) => void;

// ---------------------------------------------------------------------------
// Chunk Analysis Result
// ---------------------------------------------------------------------------

/**
 * Outcome of analyzing a single text chunk via the Mastra workflow.
 * Used to aggregate results and track partial failures across chunks.
 */
export interface ChunkResult {
  /** Zero-based index of the chunk this result corresponds to. */
  chunkIndex: number;

  /** Suggestions returned by the workflow for this chunk (empty on failure). */
  suggestions: Suggestion[];

  /** Error message if the chunk analysis failed, `undefined` on success. */
  error?: string;
}

// ---------------------------------------------------------------------------
// Change Strategy
// ---------------------------------------------------------------------------

/**
 * Type of tracked change operation, used by the Strategy pattern in the
 * Word adapter.
 *
 * All types are applied via OOXML markup (`<w:del>`, `<w:ins>`) with an
 * attached Word comment containing the justification.
 *
 * - `"insert"` — Text insertion only (`<w:ins>` markup).
 * - `"delete"` — Text deletion only (`<w:del>` markup).
 * - `"replace"` — Combined deletion + insertion (`<w:del>` + `<w:ins>`).
 */
export type ChangeType = "insert" | "delete" | "replace";

// ---------------------------------------------------------------------------
// Text Source
// ---------------------------------------------------------------------------

/**
 * Result of the text-source resolution step at the start of the analysis
 * pipeline. Encapsulates whether the text came from the user's active
 * selection or from the full document body.
 */
export interface TextSource {
  /** The plain text to analyze (selection or full document). */
  text: string;

  /**
   * `true` if the text was read from the user's current selection;
   * `false` if it was read from the document body (no active selection).
   */
  isSelection: boolean;
}

// ---------------------------------------------------------------------------
// Feedback
// ---------------------------------------------------------------------------

/**
 * Payload sent to the Mastra feedback workflow when a user accepts or rejects
 * a suggestion. Fire-and-forget — never awaited in the UI.
 *
 * Carries the target author profile slug but no user or session identifiers.
 */
export interface FeedbackPayload {
  /** Author profile slug that should receive the feedback update. */
  autorSlug: string;

  /** Editorial category label (e.g., "Redundancia"). */
  category: string;

  /** Paragraph-level context used to interpret the feedback safely. */
  context: string;

  /** Exact substring within `context` targeted by the original suggestion. */
  anchor: string;

  /**
   * The replacement text that was suggested.
   * Absent for `"comment-only"` suggestions that carry no replacement.
   */
  suggestedText?: string;

  /** Human-readable justification shown to the user. */
  justification: string;

  /** Explicit user action taken on the suggestion. */
  action: "accept" | "reject";

  /** How critical the suggestion is (from the original suggestion). */
  severity: SuggestionSeverity;

  /** Suggestion materialization kind used by the backend for interpretation. */
  suggestionType: SuggestionType;

  /** Optional free-text comment from the user (textarea). Only present when non-empty. */
  comment?: string;
}

// ---------------------------------------------------------------------------
// Profile
// ---------------------------------------------------------------------------

/**
 * An analysis profile option shown in the UI dropdown.
 */
export interface Profile {
  /** Machine-readable identifier sent to the workflow. */
  id: string;

  /** Human-readable label displayed in the dropdown. */
  label: string;
}

// ---------------------------------------------------------------------------
// Pipeline State Machine
// ---------------------------------------------------------------------------

/**
 * All possible states of the analysis pipeline.
 * Used by `PipelineStateMachine` to enforce valid transitions.
 *
 * - `idle`       — No pipeline running; ready to start.
 * - `reading`    — Reading text from the document or selection.
 * - `connecting` — Verifying backend connectivity.
 * - `chunking`   — Splitting text at paragraph boundaries.
 * - `analyzing`  — Sending chunks to the Mastra workflow.
 * - `applying`   — Applying suggestions as tracked changes.
 * - `done`       — Pipeline completed successfully.
 * - `error`      — Pipeline aborted due to unrecoverable error.
 */
export type PipelineState =
  | "idle"
  | "reading"
  | "connecting"
  | "chunking"
  | "analyzing"
  | "applying"
  | "done"
  | "error";

// ---------------------------------------------------------------------------
// Command Pattern
// ---------------------------------------------------------------------------

/**
 * Result of executing a `DocumentCommand`.
 */
export interface CommandResult {
  /** Whether the command completed successfully. */
  success: boolean;

  /** The ID of the command that produced this result. */
  commandId: string;

  /** Error message if `success` is false. */
  error?: string;

  /** Optional localized patch returned by successful real-Word apply commands. */
  mutationPatch?: ApplyMutationPatch;
}

/**
 * A reversible document operation (Command pattern).
 *
 * Commands encapsulate a single document mutation (e.g., applying one
 * suggestion as a tracked change). The `execute()` method performs the
 * operation. An `undo()` method can be added in a future iteration.
 */
export interface DocumentCommand {
  /** Stable identifier matching the source suggestion's id. */
  readonly id: string;

  /** Human-readable description for logging and UI. */
  readonly description: string;

  /** Executes the command against the Word document. */
  execute(): Promise<CommandResult>;
}

// ---------------------------------------------------------------------------
// Pipeline Result
// ---------------------------------------------------------------------------

/**
 * The aggregate outcome of a completed analysis pipeline run.
 * Produced by `ApplySuggestionsHandler` and consumed by `taskpane.ts`
 * to render the results panel.
 */
export interface PipelineResult {
  /** All unique suggestions produced by the analysis (after dedup + guard). */
  suggestions: Suggestion[];

  /** Insertion outcome (success count + failed suggestions). */
  result: InsertionResult;

  /** Error messages from chunks that failed analysis. */
  chunkErrors: string[];

  /** Whether the analysis was scoped to a text selection. */
  isSelection: boolean;

  /** Whether the pipeline was aborted before completion. */
  aborted: boolean;

  /** Human-readable reason for abortion (if `aborted` is true). */
  abortReason?: string;
}

/**
 * Visual state of a suggestion card in the taskpane after user action.
 * - "pending": Applied and awaiting user accept/reject. Buttons enabled.
 * - "resolving": User clicked; async Word API call in-flight. Buttons disabled.
 * - "accepted": User accepted from the taskpane. Terminal.
 * - "rejected": User rejected from the taskpane. Terminal.
 * - "unobservable": Word did not expose enough evidence to confirm the review
 *   state. Non-terminal — user may retry once the host state becomes visible.
 * - "identity-lost": Word exposed corrupt or incomplete v2 metadata, so the
 *   adapter cannot safely continue with compound identity semantics.
 * - "error": Word API call failed. Non-terminal — user may retry.
 */
export type SuggestionState =
  | "pending"
  | "resolving"
  | "accepted"
  | "rejected"
  | "unobservable"
  | "identity-lost"
  | "error";

/** Ordered phases emitted by the resolution workflow for observability. */
export type ResolutionPhase =
  | "locate"
  | "observe-before"
  | "execute"
  | "cleanup-comment"
  | "cleanup-anchor"
  | "inspect-after";

/** Execution summary for tracked-change resolution attempts. */
export interface ResolutionExecutionReport {
  /** Number of tracked changes the executor attempted in this workflow run. */
  attempted: number;
  /** Number of tracked changes completed before the workflow moved on or failed. */
  completed: number;
  /** Number of tracked changes still unresolved when execution stopped. */
  remaining: number;
  /** Index of the tracked change that failed, when applicable. */
  failureIndex?: number;
  /** Human-readable execution error captured at the mutation boundary. */
  error?: string;
  /**
   * Set when the executor detected that one tracked-change resolution call
   * (`accept()` / `reject()`) was a silent no-op in the host: the call
   * succeeded and `context.sync()` resolved without errors, but the document
   * tracked-change count did not decrease, indicating the proxy was stale
   * (typical for Word's `ccRange.getTrackedChanges()` proxy when the deletion
   * mark lives outside the suggestion CC range). The outer command can use
   * this signal to recover with a fresh proxy from `body.getTrackedChanges()`.
   */
  silentNoOpDetected?: {
    /** Index inside the executor's ordered step list where the no-op occurred. */
    stepIndex: number;
    /** Type of the tracked change that did not mutate the document. */
    trackedChangeType: "Added" | "Deleted";
    /** Body tracked-change count observed before the failed step. */
    bodyTrackedChangeCountBefore: number;
    /** Body tracked-change count observed after the failed step. */
    bodyTrackedChangeCountAfter: number;
  };
}

/** Structured telemetry event emitted by resolution workflows. */
export interface ResolutionTelemetryEvent {
  /** Correlation id shared by all events in one workflow attempt. */
  workflowAttemptId: string;
  /** Stable suggestion id being resolved. */
  suggestionId: string;
  /** User action requested by the taskpane. */
  action: "accept" | "reject";
  /** Resolution phase being observed. */
  phase: ResolutionPhase;
  /** Phase outcome used for later diagnostics. */
  outcome: "started" | "succeeded" | "failed" | "warning" | "reconciled";
  /** Optional structured metadata for future telemetry sinks. */
  metadata?: Record<string, string | number | boolean | null>;
}

/**
 * Result returned by `IDocumentPort.acceptSuggestion` and `IDocumentPort.rejectSuggestion`.
 */
export interface SuggestionActionResult {
  /** Final resolution status of the operation. */
  status:
    | "accepted"
    | "rejected"
    | "unobservable"
    | "identity-lost"
    | "cc-not-found"
    | "not-found"
    | "error";
  /** Number of tracked changes (insert + delete) that were accepted or rejected. */
  trackedChangesAffected: number;
  /** Whether the associated Stylistic comment was successfully deleted. */
  commentDeleted: boolean;
  /** Document-derived review state immediately after the resolution attempt. */
  pendingAfter: DocumentReviewState;
  /** Explicit document-review UI state after the resolution attempt. */
  documentState: import("./review/DocumentReviewStateMachine").DocumentReviewUiState;
  /** Human-readable error message when status is "error" or "not-found". */
  error?: string;
  /** Internal execute phase that failed, when the workflow could classify it. */
  errorPhase?: ResolutionPhase;
  /** Execution summary emitted by tracked-change resolution, when available. */
  executionReport?: ResolutionExecutionReport;
}

/** Fire-and-forget feedback dispatch semantics exposed by the resolution workflow. */
export type FeedbackDispatchStatus = "sent" | "failed" | "skipped";

/**
 * Shared resolution workflow result consumed by the taskpane.
 *
 * Extends the document mutation result with feedback observability while
 * keeping feedback non-blocking for the user.
 */
export interface SuggestionResolutionWorkflowResult
  extends SuggestionActionResult {
  /** Best-effort feedback dispatch outcome observed by the workflow. */
  feedbackStatus: FeedbackDispatchStatus;
}

/**
 * Resolution result enriched with taskpane-facing mediated review state.
 */
export interface SuggestionResolutionMediatorResult
  extends SuggestionResolutionWorkflowResult {
  /** Centralized taskpane state produced by the explicit mediator. */
  taskpaneState: ReviewTaskpaneState;
}
