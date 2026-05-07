import type {
  DocumentReviewState,
  ReviewTaskpaneState,
} from "../review/DocumentReviewStateMachine.types";
import type { SuggestionSeverity, SuggestionType } from "./Suggestion.types";

/** Payload sent to the Mastra feedback workflow. */
export interface FeedbackPayload {
  /** Author profile slug that should receive the feedback update. */
  autorSlug: string;

  /** Editorial category label. */
  category: string;

  /** Paragraph-level context used to interpret the feedback safely. */
  context: string;

  /** Exact substring within `context` targeted by the original suggestion. */
  anchor: string;

  /** Replacement text that was suggested. */
  suggestedText?: string;

  /** Human-readable justification shown to the user. */
  justification: string;

  /** Explicit user action taken on the suggestion. */
  action: "accept" | "reject";

  /** How critical the suggestion is. */
  severity: SuggestionSeverity;

  /** Suggestion materialization kind used by the backend for interpretation. */
  suggestionType: SuggestionType;

  /** Optional free-text comment from the user. */
  comment?: string;
}

/** Ordered phases emitted by the resolution workflow for observability. */
export type ResolutionPhase =
  | "locate"
  | "observe-before"
  | "execute"
  | "cleanup-comment"
  | "cleanup-metadata"
  | "cleanup-anchor"
  | "inspect-after";

/** Execution summary for tracked-change resolution attempts. */
export interface ResolutionExecutionReport {
  /** Number of tracked changes the executor attempted. */
  attempted: number;
  /** Number of tracked changes completed. */
  completed: number;
  /** Number of tracked changes still unresolved when execution stopped. */
  remaining: number;
  /** Index of the tracked change that failed, when applicable. */
  failureIndex?: number;
  /** Human-readable execution error captured at the mutation boundary. */
  error?: string;
  /** Host evidence for a detected silent no-op. */
  silentNoOpDetected?: {
    stepIndex: number;
    trackedChangeType: "Added" | "Deleted";
    bodyTrackedChangeCountBefore: number;
    bodyTrackedChangeCountAfter: number;
  };
  /** Host evidence for a mutation that could not be semantically verified. */
  unverifiedMutation?: {
    stepIndex: number;
    trackedChangeType: "Added" | "Deleted";
    reason: "body-count-probe-failed";
    bodyTrackedChangeCountBefore?: number;
    bodyTrackedChangeCountAfter?: number;
    bodyTrackedChangeCountBeforeError?: string;
    bodyTrackedChangeCountAfterError?: string;
  };
}

/** Scalar metadata carried by best-effort resolution observability records. */
export type ResolutionObservabilityMetadata = Record<string, string | number | boolean | null>;

/** Structured phase event emitted by resolution workflows. */
export interface ResolutionObservabilityEvent {
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
  /** Optional structured metadata for future observability sinks. */
  metadata?: ResolutionObservabilityMetadata;
}

/** Snapshot checkpoints captured around risky resolution transitions. */
export type ResolutionObservabilitySnapshotLabel =
  | "after-execute-before-cleanup"
  | "after-cleanup-before-return";

/** Host evidence snapshot emitted for forensic debugging of resolution workflows. */
export interface ResolutionObservabilitySnapshot {
  /** Correlation id shared by all records in one workflow attempt. */
  workflowAttemptId: string;
  /** Stable suggestion id being resolved. */
  suggestionId: string;
  /** User action requested by the taskpane. */
  action: "accept" | "reject";
  /** Workflow checkpoint where the snapshot was captured. */
  label: ResolutionObservabilitySnapshotLabel;
  /** Structured workflow-owned state known at snapshot time. */
  workflowState: Record<string, unknown>;
  /** Optional host-facing evidence gathered from Word during the snapshot. */
  hostEvidence?: Record<string, unknown>;
  /** Optional heuristic or debug-only details associated with the snapshot. */
  heuristicDiagnostics?: Record<string, unknown>;
}

/** Result returned by document-port accept/reject operations. */
export interface SuggestionActionResult {
  /** Final resolution status of the operation. */
  status:
    | "accepted"
    | "rejected"
    | "unobservable"
    | "identity-lost"
    | "ambiguous-location"
    | "mixed-group"
    | "cc-not-found"
    | "not-found"
    | "error";
  /** Number of tracked changes that were accepted or rejected. */
  trackedChangesAffected: number;
  /** Whether the associated Stylistic comment was successfully deleted. */
  commentDeleted: boolean;
  /** Document-derived review state immediately after the resolution attempt. */
  pendingAfter: DocumentReviewState;
  /** Explicit document-review UI state after the resolution attempt. */
  documentState: ReviewTaskpaneState["documentState"];
  /** Human-readable error message when status is terminal failure. */
  error?: string;
  /** Internal execute phase that failed, when classifiable. */
  errorPhase?: ResolutionPhase;
  /** Execution summary emitted by tracked-change resolution, when available. */
  executionReport?: ResolutionExecutionReport;
}

/** Fire-and-forget feedback dispatch semantics exposed by the resolution workflow. */
export type FeedbackDispatchStatus = "sent" | "failed" | "skipped";

/** Shared resolution workflow result consumed by the taskpane. */
export interface SuggestionResolutionWorkflowResult extends SuggestionActionResult {
  /** Best-effort feedback dispatch outcome observed by the workflow. */
  feedbackStatus: FeedbackDispatchStatus;
}

/** Resolution result enriched with taskpane-facing mediated review state. */
export interface SuggestionResolutionMediatorResult extends SuggestionResolutionWorkflowResult {
  /** Centralized taskpane state produced by the explicit mediator. */
  taskpaneState: ReviewTaskpaneState;
}
/** User-driven suggestion resolution action. */
export type ResolutionAction = "accept" | "reject";
