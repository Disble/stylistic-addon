/* global Word */

import type { IResolutionObservabilityPort } from "../../../domain/ports";
import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import type { SuggestionActionResult } from "../../../domain/suggestion/SuggestionResolutionWorkflow.types";
import { NoopResolutionObservabilityAdapter } from "../../observability/NoopResolutionObservabilityAdapter";
import { getDefaultTextLocator, type TextLocator } from "../WordTextLocatorContext";
import { CommentOnlySuggestionResolver } from "./CommentOnlySuggestionResolver";
import { DocumentReviewStateInspector } from "./DocumentReviewStateInspector";
import type { ResolveSuggestionOutcome } from "./ResolutionContext";
import { ResolutionErrorSerializer } from "./ResolutionErrorParser";
import { ResolutionObservabilityReporter } from "./ResolutionObservabilityAdapter";
import { ResolveSuggestionCommentOnlyOrchestrator } from "./ResolveSuggestionCommentOnlyOrchestrator";
import { ResolveSuggestionResultFactory } from "./ResolveSuggestionResultFactory";
import { ResolveSuggestionTrackChangeOrchestrator } from "./ResolveSuggestionTrackChangeOrchestrator";
import { SuggestionLocator } from "./SuggestionLocator";
import { SuggestionResolutionCleanup } from "./SuggestionResolutionCleanup";
import { SuggestionResolutionObserver } from "./SuggestionResolutionObserver";

/**
 * Resolves one suggestion using the operational wrapper as the only mutation scope.
 *
 * The workflow intentionally avoids reconstructing replace semantics from partial
 * host evidence (`cc`, `ccRange`, `comment`, etc.). For track-change suggestions
 * it mutates only the tracked-change collection exposed by the wrapper range. If
 * that scope is missing or not actionable, the command fails closed before any
 * mutation.
 */
export class ResolveSuggestionCommand {
  private readonly stateInspector: DocumentReviewStateInspector;
  private readonly resultFactory: ResolveSuggestionResultFactory;
  private readonly observabilityReporter: ResolutionObservabilityReporter;
  private readonly commentOnlyOrchestrator: ResolveSuggestionCommentOnlyOrchestrator;
  private readonly trackChangeOrchestrator: ResolveSuggestionTrackChangeOrchestrator;
  private readonly errorSerializer = new ResolutionErrorSerializer();
  private workflowAttemptId = "";

  constructor(
    private readonly suggestion: Suggestion,
    private readonly action: "accept" | "reject",
    textLocator: TextLocator = getDefaultTextLocator(),
    observabilityPort: IResolutionObservabilityPort = new NoopResolutionObservabilityAdapter()
  ) {
    this.stateInspector = new DocumentReviewStateInspector();
    const locator = new SuggestionLocator(suggestion);
    const cleanup = new SuggestionResolutionCleanup(suggestion.id, action);
    this.resultFactory = new ResolveSuggestionResultFactory(action, this.stateInspector);
    const commentOnlyResolver = new CommentOnlySuggestionResolver(
      suggestion.id,
      this.resultFactory,
      this.stateInspector
    );
    const observer = new SuggestionResolutionObserver(suggestion, locator, textLocator);
    this.observabilityReporter = new ResolutionObservabilityReporter(
      suggestion.id,
      action,
      observabilityPort
    );
    this.commentOnlyOrchestrator = new ResolveSuggestionCommentOnlyOrchestrator(
      locator,
      cleanup,
      commentOnlyResolver,
      this.resultFactory,
      this.stateInspector,
      this.observabilityReporter
    );
    this.trackChangeOrchestrator = new ResolveSuggestionTrackChangeOrchestrator(
      suggestion,
      action,
      locator,
      cleanup,
      observer,
      this.resultFactory,
      this.stateInspector,
      this.observabilityReporter,
      this.errorSerializer
    );
  }

  /** Executes the command and returns a stable result instead of throwing. */
  async execute(): Promise<SuggestionActionResult> {
    try {
      return await Word.run(async (context) => {
        const outcome = await this.executeWithinContext(context);
        return this.resultFactory.buildResolutionResult(
          outcome.status,
          outcome.trackedChangesAffected,
          outcome.commentDeleted,
          outcome.pendingBefore,
          outcome.pendingAfter,
          outcome.error,
          outcome.executionReport
        );
      });
    } catch (error) {
      const serialized = this.errorSerializer.serialize(error);
      const pendingAfter = await Word.run((context) => this.stateInspector.inspect(context)).catch(
        () => this.stateInspector.buildEmptyState()
      );

      return this.resultFactory.buildErrorResult(serialized.message, pendingAfter);
    }
  }

  /** Runs the resolution sequence inside one `Word.run` context. */
  private async executeWithinContext(
    context: Word.RequestContext
  ): Promise<ResolveSuggestionOutcome> {
    this.workflowAttemptId = `${this.suggestion.id}:${this.action}:${Date.now()}`;
    this.observabilityReporter.setWorkflowAttemptId(this.workflowAttemptId);

    if (this.suggestion.type === "comment-only") {
      return this.commentOnlyOrchestrator.execute(context);
    }

    return this.trackChangeOrchestrator.execute(context);
  }
}
