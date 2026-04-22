/* global Word, console */

/**
 * ResolveSuggestionCommand — Command pattern for accepting/rejecting suggestions.
 *
 * Encapsulates the complete logic for resolving a single `Suggestion` in Word:
 * finding the Content Control, observing tracked changes through multiple
 * evidence sources (CC-scoped, CC-range, body-level, operational anchor,
 * colocated comment range), applying the terminal action, and cleaning up
 * artifacts.
 *
 * Parallel to `ApplySuggestionCommand` which handles *applying* suggestions.
 * This command handles *resolving* them after the user accepts or rejects.
 *
 * Never throws — catches all errors and returns a `SuggestionActionResult`.
 *
 * @module ResolveSuggestionCommand
 */

import type {
  ResolutionExecutionReport,
  Suggestion,
  SuggestionActionResult,
  SuggestionResolutionWarning,
} from "../../domain/types";
import { CommentOnlySuggestionResolver } from "./resolution/CommentOnlySuggestionResolver";
import { DocumentReviewStateInspector } from "./resolution/DocumentReviewStateInspector";
import { ResolveSuggestionResultFactory } from "./resolution/ResolveSuggestionResultFactory";
import { SuggestionLocator } from "./resolution/SuggestionLocator";
import { SuggestionResolutionCleanup } from "./resolution/SuggestionResolutionCleanup";
import { SuggestionResolutionObserver } from "./resolution/SuggestionResolutionObserver";
import { TrackedChangeResolutionExecutor } from "./resolution/TrackedChangeResolutionExecutor";
import {
  getDefaultTextLocator,
  type TextLocator,
} from "./WordTextLocatorContext";

// ---------------------------------------------------------------------------
// ResolveSuggestionCommand
// ---------------------------------------------------------------------------

/**
 * Command that resolves (accepts or rejects) a single suggestion in Word.
 *
 * Usage:
 * ```ts
 * const result = await new ResolveSuggestionCommand(suggestion, "accept").execute();
 * ```
 */
export class ResolveSuggestionCommand {
  private readonly stateInspector: DocumentReviewStateInspector;
  private readonly locator: SuggestionLocator;
  private readonly cleanup: SuggestionResolutionCleanup;
  private readonly executor: TrackedChangeResolutionExecutor;
  private readonly resultFactory: ResolveSuggestionResultFactory;
  private readonly commentOnlyResolver: CommentOnlySuggestionResolver;
  private readonly observer: SuggestionResolutionObserver;
  private lastExecutionReport?: ResolutionExecutionReport;

  constructor(
    private readonly suggestion: Suggestion,
    private readonly action: "accept" | "reject",
    textLocator: TextLocator = getDefaultTextLocator(),
  ) {
    this.stateInspector = new DocumentReviewStateInspector();
    this.locator = new SuggestionLocator(suggestion);
    this.cleanup = new SuggestionResolutionCleanup(suggestion.id, action);
    this.executor = new TrackedChangeResolutionExecutor(suggestion.id, action);
    this.resultFactory = new ResolveSuggestionResultFactory(
      action,
      this.stateInspector,
    );
    this.commentOnlyResolver = new CommentOnlySuggestionResolver(
      suggestion.id,
      this.resultFactory,
      this.stateInspector,
    );
    this.observer = new SuggestionResolutionObserver(
      suggestion,
      this.locator,
      textLocator,
    );
  }

  /**
   * Executes the resolution command.
   * Never throws — catches all errors and returns a result object.
   */
  async execute(): Promise<SuggestionActionResult> {
    try {
      return await Word.run(async (context) => {
        this.lastExecutionReport = undefined;
        console.log(
          `🎯 [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" type=${this.suggestion.type}`,
        );
        const pendingBefore = await this.stateInspector.inspect(context);
        const { rankedCandidates, selectedCc: cc } =
          await this.locator.locateResolutionArtifacts(context);

        if (!cc) {
          console.warn(
            `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" failed: CC not found`,
          );
          return this.resultFactory.buildResolutionResult(
            "cc-not-found",
            0,
            false,
            pendingBefore,
            pendingBefore,
          );
        }

        console.log(
          `🎯 [ResolveSuggestionCommand] selected CC for suggestionId="${this.suggestion.id}": tag="${cc.tag}"`,
        );

        if (this.suggestion.type === "comment-only") {
          const colocatedComment =
            await this.locator.findColocatedStylisticComment(context, cc);
          const commentDeleted =
            await this.cleanup.deleteLocatedStylisticComment(
              context,
              colocatedComment,
            );
          return this.commentOnlyResolver.resolve({
            context,
            cc,
            commentDeleted,
            pendingBefore,
          });
        }

        const observation = await this.observer.observeResolutionCandidates(
          context,
          rankedCandidates,
          cc,
        );

        if (observation.observationStatus === "identity-lost") {
          console.warn(
            `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" ended in identity-lost`,
          );
          return this.resultFactory.buildObservationFailureResult(
            context,
            "identity-lost",
            pendingBefore,
          );
        }

        if (
          observation.observationStatus !== "confirmed-pending" ||
          observation.trackedChanges.length === 0
        ) {
          console.warn(
            `⚠️ [ResolveSuggestionCommand] action=${this.action} suggestionId="${this.suggestion.id}" ended in unobservable after all evidence sources`,
          );
          return this.resultFactory.buildObservationFailureResult(
            context,
            "unobservable",
            pendingBefore,
          );
        }

        const executionReport = this.executor.apply(observation.trackedChanges);
        this.lastExecutionReport = executionReport;

        if (executionReport.error) {
          throw new Error(executionReport.error);
        }

        const warnings: SuggestionResolutionWarning[] = [];
        const commentCleanup =
          await this.cleanup.deleteLocatedStylisticCommentAfterResolution(
            context,
            observation.selectedComment,
          );
        if (commentCleanup.warning) {
          warnings.push(commentCleanup.warning);
        }

        const anchorCleanupWarning =
          await this.cleanup.cleanupResolvedSuggestionAnchor(
            context,
            observation.selectedCc,
          );
        if (anchorCleanupWarning) {
          warnings.push(anchorCleanupWarning);
        }

        const inspectAfter = await this.stateInspector.inspectAfterResolution(
          context,
          pendingBefore,
          this.action,
          this.suggestion.id,
        );
        if (inspectAfter.warning) {
          warnings.push(inspectAfter.warning);
        }

        return this.resultFactory.buildResolutionResult(
          this.resultFactory.toResolutionStatus(),
          executionReport.completed,
          commentCleanup.deleted,
          pendingBefore,
          inspectAfter.pendingAfter,
          undefined,
          warnings,
          executionReport,
        );
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [ResolveSuggestionCommand] threw for suggestionId="${this.suggestion.id}": ${message}`,
      );
      const pendingAfter = await Word.run((ctx) =>
        this.stateInspector.inspect(ctx),
      ).catch(() => this.stateInspector.buildEmptyState());

      const reconciliation = await this.tryBuildSemanticReconciliationResult(
        pendingAfter,
        message,
      );
      if (reconciliation) {
        return reconciliation;
      }

      return this.resultFactory.buildErrorResult(
        message,
        pendingAfter,
        undefined,
        this.lastExecutionReport,
      );
    }
  }

  /** Reconciles late failures against the semantic execution report. */
  private async tryBuildSemanticReconciliationResult(
    pendingAfter: import("../../domain/types").DocumentReviewState,
    message: string,
  ): Promise<SuggestionActionResult | null> {
    if (!this.lastExecutionReport || this.lastExecutionReport.remaining > 0) {
      return null;
    }

    return this.resultFactory.buildResolutionResult(
      this.resultFactory.toResolutionStatus(),
      this.lastExecutionReport.completed,
      false,
      pendingAfter,
      pendingAfter,
      undefined,
      [
        {
          code: "cleanup-failed",
          phase: "cleanup",
          message,
        },
        {
          code: "reconciled-after-error",
          phase: "reconcile",
          message,
        },
      ],
      this.lastExecutionReport,
    );
  }
}
