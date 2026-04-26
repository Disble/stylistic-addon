import type {
  DocumentReviewState,
  DocumentReviewUiState,
} from "./review/DocumentReviewStateMachine.types";
import type { Suggestion } from "./suggestion/Suggestion.types";

/** Machine-readable failure reasons for one suggestion application attempt. */
export type SuggestionApplicationFailureReason =
  | "not-found"
  | "covered-by-existing-cc"
  | "command-error";

/** One failed suggestion application with preserved reason and message. */
export interface SuggestionApplicationFailure {
  /** The suggestion that could not be applied. */
  suggestion: Suggestion;

  /** Stable failure reason used by UI and telemetry. */
  reason: SuggestionApplicationFailureReason;

  /** Human-readable adapter message captured at the failure boundary. */
  message: string;
}

/** Result of attempting to insert suggestions as tracked changes in Word. */
export interface InsertionResult {
  /** Number of suggestions successfully applied as tracked changes. */
  successCount: number;

  /** Suggestions that failed to apply, with preserved failure semantics. */
  failedSuggestions: SuggestionApplicationFailure[];
}

/** Batch insertion result enriched with document-review semantics. */
export interface ApplySuggestionsResult extends InsertionResult {
  /** Document-derived review state after the batch finishes. */
  pendingAfter: DocumentReviewState;

  /** Explicit document-review UI state after the batch finishes. */
  documentState: DocumentReviewUiState;

  /** Whether this batch had to activate Track Changes lazily. */
  trackChangesActivatedForBatch: boolean;
}

/** Localized mutation patch captured after a successful real-Word apply. */
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

/** Type of tracked change operation used by the Word adapter strategy. */
export type ChangeType = "insert" | "delete" | "replace";

/** Result of executing a `DocumentCommand`. */
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

/** A reversible document operation (Command pattern). */
export interface DocumentCommand {
  /** Stable identifier matching the source suggestion's id. */
  readonly id: string;

  /** Human-readable description for logging and UI. */
  readonly description: string;

  /** Executes the command against the Word document. */
  execute(): Promise<CommandResult>;
}
