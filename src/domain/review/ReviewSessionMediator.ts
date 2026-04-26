/**
 * Review Session Mediator.
 *
 * Centralizes multi-party coordination for the taskpane review experience:
 * - document mutation via `SuggestionResolutionWorkflow`
 * - document-review UI state via `DocumentReviewStateMachine`
 * - cleanup section visibility semantics via `IDocumentPort`
 *
 * This is an explicit Mediator, not a passive event broadcaster.
 * Its purpose is to own cross-component policy so `taskpane.ts` stops combining
 * adapter results and UI booleans procedurally.
 *
 * @module ReviewSessionMediator
 */

import type { IDocumentPort, IFeedbackPort } from "../ports";
import type { Suggestion } from "../suggestion/Suggestion.types";
import { SuggestionResolutionWorkflow } from "../suggestion/SuggestionResolutionWorkflow";
import type { SuggestionResolutionMediatorResult } from "../suggestion/SuggestionResolutionWorkflow.types";
import { DocumentReviewStateMachine } from "./DocumentReviewStateMachine";
import type { ReviewTaskpaneState } from "./DocumentReviewStateMachine.types";

type ResolutionAction = "accept" | "reject";

/**
 * Explicit mediator for taskpane review coordination.
 */
export class ReviewSessionMediator {
  private readonly documentReviewStateMachine =
    new DocumentReviewStateMachine();

  constructor(
    private readonly documentPort: IDocumentPort,
    feedbackPort: IFeedbackPort,
  ) {
    this.resolutionWorkflow = new SuggestionResolutionWorkflow(
      documentPort,
      feedbackPort,
    );
  }

  private readonly resolutionWorkflow: SuggestionResolutionWorkflow;

  /**
   * Rehydrates the taskpane review state from the authoritative document.
   */
  async rehydrateTaskpaneState(): Promise<ReviewTaskpaneState> {
    const reviewState = await this.documentPort.getDocumentReviewState();
    this.documentReviewStateMachine.syncFromDocument(reviewState);
    const cleanup = await this.documentPort.getCleanupPreview();
    return this.buildTaskpaneState(cleanup.deletable > 0);
  }

  /**
   * Accepts one suggestion and returns the centralized taskpane state.
   */
  async acceptSuggestion(
    suggestion: Suggestion,
    comment?: string,
  ): Promise<SuggestionResolutionMediatorResult> {
    return this.resolveSuggestion(suggestion, "accept", comment);
  }

  /**
   * Rejects one suggestion and returns the centralized taskpane state.
   */
  async rejectSuggestion(
    suggestion: Suggestion,
    comment?: string,
  ): Promise<SuggestionResolutionMediatorResult> {
    return this.resolveSuggestion(suggestion, "reject", comment);
  }

  /**
   * Applies the explicit user action of disabling Track Changes and returns the
   * resulting centralized taskpane state.
   */
  async disableTrackChanges(): Promise<ReviewTaskpaneState> {
    await this.documentPort.disableTrackChanges();
    this.documentReviewStateMachine.disableTrackChanges();
    const cleanup = await this.documentPort.getCleanupPreview();
    return this.buildTaskpaneState(cleanup.deletable > 0);
  }

  /**
   * Shared resolution branch coordinated through the workflow + state machine.
   */
  private async resolveSuggestion(
    suggestion: Suggestion,
    action: ResolutionAction,
    comment?: string,
  ): Promise<SuggestionResolutionMediatorResult> {
    const workflowResult =
      action === "accept"
        ? await this.resolutionWorkflow.acceptSuggestion(suggestion, comment)
        : await this.resolutionWorkflow.rejectSuggestion(suggestion, comment);

    this.documentReviewStateMachine.syncFromDocument(
      workflowResult.pendingAfter,
    );

    const cleanup = await this.documentPort.getCleanupPreview();

    return {
      ...workflowResult,
      taskpaneState: this.buildTaskpaneState(cleanup.deletable > 0),
    };
  }

  /**
   * Builds the taskpane-facing state from the explicit review machine + cleanup.
   */
  private buildTaskpaneState(showCleanupSection: boolean): ReviewTaskpaneState {
    return {
      documentState: this.documentReviewStateMachine.state,
      showDisableTrackChangesCta:
        this.documentReviewStateMachine.shouldShowDisableTrackChangesCta,
      showCleanupSection,
    };
  }
}
