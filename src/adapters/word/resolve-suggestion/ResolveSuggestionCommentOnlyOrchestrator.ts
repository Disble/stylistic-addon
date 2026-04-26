import type { DocumentReviewState } from "../../../domain/review/DocumentReviewStateMachine.types";
import type { SuggestionActionResult } from "../../../domain/suggestion/SuggestionResolutionWorkflow.types";
import type { CommentOnlySuggestionResolver } from "./CommentOnlySuggestionResolver";
import type { DocumentReviewStateInspector } from "./DocumentReviewStateInspector";
import type { ResolveSuggestionOutcome } from "./ResolutionContext";
import type { ResolutionObservabilityReporter } from "./ResolutionObservabilityAdapter";
import type { ResolveSuggestionResultFactory } from "./ResolveSuggestionResultFactory";
import type { SuggestionLocator } from "./SuggestionLocator";
import type { SuggestionResolutionCleanup } from "./SuggestionResolutionCleanup";

/**
 * Orchestrates the dedicated comment-only resolution branch.
 *
 * This workflow intentionally stays outside the tracked-change path: locate the
 * canonical comment-only anchor, delete the colocated Stylistic comment when
 * present, delete the anchor CC, and return one stable adapter result.
 */
export class ResolveSuggestionCommentOnlyOrchestrator {
  constructor(
    private readonly locator: SuggestionLocator,
    private readonly cleanup: SuggestionResolutionCleanup,
    private readonly resolver: CommentOnlySuggestionResolver,
    private readonly resultFactory: ResolveSuggestionResultFactory,
    private readonly stateInspector: DocumentReviewStateInspector,
    private readonly observabilityReporter: ResolutionObservabilityReporter,
  ) {}

  /** Runs the full comment-only workflow inside the active Word batch. */
  async execute(
    context: Word.RequestContext,
  ): Promise<ResolveSuggestionOutcome> {
    await this.observabilityReporter.emitPhase("locate", "started", {
      suggestionType: "comment-only",
    });

    const { candidates, selectedCc, locateStatus } =
      await this.locator.locateCommentOnlyArtifacts(context);
    const pendingBefore = await this.stateInspector.inspect(context);

    await this.observabilityReporter.emitPhase(
      "locate",
      selectedCc ? "succeeded" : "failed",
      {
        candidateCount: candidates.length,
        selectedCcFound: Boolean(selectedCc),
        locateStatus,
      },
    );

    if (locateStatus === "ambiguous-location") {
      const result = await this.resultFactory.buildObservationFailureResult(
        context,
        locateStatus,
        pendingBefore,
      );
      return this.toOutcome(result, pendingBefore);
    }

    if (!selectedCc) {
      return {
        status: "cc-not-found",
        trackedChangesAffected: 0,
        commentDeleted: false,
        pendingBefore,
        pendingAfter: pendingBefore,
      };
    }

    const colocatedComment = await this.locator.findColocatedStylisticComment(
      context,
      selectedCc,
    );
    const commentDeleted = await this.cleanup.deleteLocatedStylisticComment(
      context,
      colocatedComment,
    );
    const result = await this.resolver.resolve({
      context,
      cc: selectedCc,
      commentDeleted,
      pendingBefore,
    });

    return this.toOutcome(result, pendingBefore);
  }

  /** Keeps the command-facing internal outcome shape stable. */
  private toOutcome(
    result: SuggestionActionResult,
    pendingBefore: DocumentReviewState,
  ): ResolveSuggestionOutcome {
    return {
      status: result.status,
      trackedChangesAffected: result.trackedChangesAffected,
      commentDeleted: result.commentDeleted,
      pendingBefore,
      pendingAfter: result.pendingAfter,
      error: result.error,
      executionReport: result.executionReport,
    };
  }
}
