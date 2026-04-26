/* global console */

import type { IResolutionObservabilityPort } from "../../../domain/ports";
import type {
  ResolutionObservabilityEvent,
  ResolutionObservabilityMetadata,
  ResolutionObservabilitySnapshot,
  ResolutionObservabilitySnapshotLabel,
  ResolutionPhase,
} from "../../../domain/suggestion/SuggestionResolutionWorkflow.types";

/** Encapsulates best-effort resolution observability emission for one suggestion workflow. */
export class ResolutionObservabilityReporter {
  private workflowAttemptId = "";

  constructor(
    private readonly suggestionId: string,
    private readonly action: "accept" | "reject",
    private readonly observabilityPort: IResolutionObservabilityPort,
  ) {}

  /** Updates the correlation id used by subsequent records in the same workflow run. */
  setWorkflowAttemptId(workflowAttemptId: string): void {
    this.workflowAttemptId = workflowAttemptId;
  }

  /** Merges two optional metadata objects without leaking observability assembly details into callers. */
  mergeMetadata(
    base?: ResolutionObservabilityMetadata,
    extra?: ResolutionObservabilityMetadata,
  ): ResolutionObservabilityMetadata {
    if (!base) {
      return extra ?? {};
    }

    if (!extra) {
      return base;
    }

    return {
      ...base,
      ...extra,
    };
  }

  /** Emits one structured phase event and degrades adapter failures to warnings. */
  async emitPhase(
    phase: ResolutionPhase,
    outcome: ResolutionObservabilityEvent["outcome"],
    metadata?: ResolutionObservabilityEvent["metadata"],
  ): Promise<void> {
    try {
      await this.observabilityPort.emitEvent({
        workflowAttemptId: this.workflowAttemptId,
        suggestionId: this.suggestionId,
        action: this.action,
        phase,
        outcome,
        metadata,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [ResolveSuggestionCommand] observability event failed for suggestionId="${this.suggestionId}" phase=${phase}: ${message}`,
      );
    }
  }

  /** Captures one structured host-evidence snapshot and degrades adapter failures to warnings. */
  async captureSnapshot(
    label: ResolutionObservabilitySnapshotLabel,
    workflowState: ResolutionObservabilitySnapshot["workflowState"],
    hostEvidence?: ResolutionObservabilitySnapshot["hostEvidence"],
    heuristicDiagnostics?: ResolutionObservabilitySnapshot["heuristicDiagnostics"],
  ): Promise<void> {
    try {
      await this.observabilityPort.captureSnapshot({
        workflowAttemptId: this.workflowAttemptId,
        suggestionId: this.suggestionId,
        action: this.action,
        label,
        workflowState,
        ...(hostEvidence ? { hostEvidence } : {}),
        ...(heuristicDiagnostics ? { heuristicDiagnostics } : {}),
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [ResolveSuggestionCommand] observability snapshot failed for suggestionId="${this.suggestionId}" label=${label}: ${message}`,
      );
    }
  }
}
