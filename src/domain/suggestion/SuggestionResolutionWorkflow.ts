/**
 * Suggestion resolution workflow.
 *
 * Owns the shared application flow for accept/reject interactions:
 * 1. resolve the suggestion in Word through `IDocumentPort`
 * 2. derive pending/CTA semantics from the document result
 * 3. send feedback best-effort without blocking the UI
 * 4. return an observable result for the taskpane
 *
 * The document remains the source of truth for pending Stylistic artifacts.
 */

import type { IDocumentPort, IFeedbackPort } from "../ports";
import type { Suggestion } from "./Suggestion.types";
import type {
  FeedbackDispatchStatus,
  FeedbackPayload,
  ResolutionAction,
  SuggestionActionResult,
  SuggestionResolutionWorkflowResult,
} from "./SuggestionResolutionWorkflow.types";

/**
 * Workflow-oriented façade for suggestion resolution semantics.
 */
export class SuggestionResolutionWorkflow {
  constructor(
    private readonly documentPort: IDocumentPort,
    private readonly feedbackPort: IFeedbackPort
  ) {}

  /** Resolves one suggestion as accepted. */
  async acceptSuggestion(
    suggestion: Suggestion,
    comment?: string
  ): Promise<SuggestionResolutionWorkflowResult> {
    return this.resolveSuggestion(suggestion, "accept", comment);
  }

  /** Resolves one suggestion as rejected. */
  async rejectSuggestion(
    suggestion: Suggestion,
    comment?: string
  ): Promise<SuggestionResolutionWorkflowResult> {
    return this.resolveSuggestion(suggestion, "reject", comment);
  }

  /**
   * Shared resolution sequence with Strategy-style action variance.
   */
  private async resolveSuggestion(
    suggestion: Suggestion,
    action: ResolutionAction,
    comment?: string
  ): Promise<SuggestionResolutionWorkflowResult> {
    const documentResult =
      action === "accept"
        ? await this.documentPort.acceptSuggestion(suggestion)
        : await this.documentPort.rejectSuggestion(suggestion);

    const feedbackStatus = await this.dispatchFeedback(suggestion, action, documentResult, comment);

    return {
      ...documentResult,
      feedbackStatus,
    };
  }

  /**
   * Sends feedback best-effort without delaying the resolution outcome.
   */
  private async dispatchFeedback(
    suggestion: Suggestion,
    action: ResolutionAction,
    result: SuggestionActionResult,
    comment?: string
  ): Promise<FeedbackDispatchStatus> {
    if (!this.shouldSendFeedback(result)) {
      return "skipped";
    }

    try {
      const payload = await this.buildFeedbackPayload(suggestion, action, comment);
      const feedbackPromise = this.feedbackPort.sendFeedback(payload);
      void feedbackPromise.catch(() => undefined);
      return "sent";
    } catch {
      return "failed";
    }
  }

  /**
   * Feedback is only emitted for explicit terminal user actions.
   * Ambiguous observation states such as `unobservable` and `identity-lost`
   * must never emit feedback.
   */
  private shouldSendFeedback(result: SuggestionActionResult): boolean {
    return result.status === "accepted" || result.status === "rejected";
  }

  /**
   * Creates the feedback payload that mirrors the user's explicit action.
   */
  private async buildFeedbackPayload(
    suggestion: Suggestion,
    action: ResolutionAction,
    comment?: string
  ): Promise<FeedbackPayload> {
    const trimmedComment = comment?.trim();
    const documentUuid = await this.documentPort.getDocumentUuid();

    return {
      documentUuid,
      category: suggestion.category,
      context: suggestion.context,
      anchor: suggestion.anchor,
      ...(suggestion.suggestedText === undefined
        ? {}
        : { suggestedText: suggestion.suggestedText }),
      justification: suggestion.justification,
      action,
      severity: suggestion.severity,
      suggestionType: suggestion.type,
      ...(trimmedComment ? { comment: trimmedComment } : {}),
    };
  }
}
