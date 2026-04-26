/* global Word, console */

import type { IResolutionObservabilityPort } from "../../domain/ports";
import type { DocumentReviewState } from "../../domain/review/DocumentReviewStateMachine.types";
import type { Suggestion } from "../../domain/suggestion/Suggestion.types";
import type {
  ResolutionExecutionReport,
  SuggestionActionResult,
} from "../../domain/suggestion/SuggestionResolutionWorkflow.types";
import { NoopResolutionObservabilityAdapter } from "../observability/NoopResolutionObservabilityAdapter";
import { CommentOnlySuggestionResolver } from "./resolution/CommentOnlySuggestionResolver";
import { DocumentReviewStateInspector } from "./resolution/DocumentReviewStateInspector";
import type { ResolutionObservation } from "./resolution/ResolutionContext";
import { ResolutionErrorSerializer } from "./resolution/ResolutionErrorParser";
import { ResolutionObservabilityReporter } from "./resolution/ResolutionObservabilityAdapter";
import { ResolveSuggestionResultFactory } from "./resolution/ResolveSuggestionResultFactory";
import { SuggestionLocator } from "./resolution/SuggestionLocator";
import { SuggestionResolutionCleanup } from "./resolution/SuggestionResolutionCleanup";
import { SuggestionResolutionObserver } from "./resolution/SuggestionResolutionObserver";
import {
  getDefaultTextLocator,
  type TextLocator,
} from "./WordTextLocatorContext";

type CohesiveResolutionOutcome = {
  status: SuggestionActionResult["status"];
  trackedChangesAffected: number;
  commentDeleted: boolean;
  pendingBefore: DocumentReviewState;
  pendingAfter: DocumentReviewState;
  executionReport?: ResolutionExecutionReport;
  error?: string;
};

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
  private readonly locator: SuggestionLocator;
  private readonly cleanup: SuggestionResolutionCleanup;
  private readonly resultFactory: ResolveSuggestionResultFactory;
  private readonly commentOnlyResolver: CommentOnlySuggestionResolver;
  private readonly observer: SuggestionResolutionObserver;
  private readonly observabilityReporter: ResolutionObservabilityReporter;
  private readonly errorSerializer = new ResolutionErrorSerializer();
  private workflowAttemptId = "";

  constructor(
    private readonly suggestion: Suggestion,
    private readonly action: "accept" | "reject",
    textLocator: TextLocator = getDefaultTextLocator(),
    observabilityPort: IResolutionObservabilityPort = new NoopResolutionObservabilityAdapter(),
  ) {
    this.stateInspector = new DocumentReviewStateInspector();
    this.locator = new SuggestionLocator(suggestion);
    this.cleanup = new SuggestionResolutionCleanup(suggestion.id, action);
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
    this.observabilityReporter = new ResolutionObservabilityReporter(
      suggestion.id,
      action,
      observabilityPort,
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
          outcome.executionReport,
        );
      });
    } catch (error) {
      const serialized = this.errorSerializer.serialize(error);
      const pendingAfter = await Word.run((context) =>
        this.stateInspector.inspect(context),
      ).catch(() => this.stateInspector.buildEmptyState());

      return this.resultFactory.buildErrorResult(
        serialized.message,
        pendingAfter,
      );
    }
  }

  /** Runs the resolution sequence inside one `Word.run` context. */
  private async executeWithinContext(
    context: Word.RequestContext,
  ): Promise<CohesiveResolutionOutcome> {
    this.workflowAttemptId = `${this.suggestion.id}:${this.action}:${Date.now()}`;
    this.observabilityReporter.setWorkflowAttemptId(this.workflowAttemptId);

    await this.observabilityReporter.emitPhase("locate", "started", {
      suggestionType: this.suggestion.type,
    });

    const { candidates, selectedCc, locateStatus } =
      await this.locator.locateResolutionArtifacts(context);
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

    if (
      locateStatus === "ambiguous-location" ||
      locateStatus === "identity-lost"
    ) {
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

    if (this.suggestion.type === "comment-only") {
      return this.resolveCommentOnly(context, selectedCc, pendingBefore);
    }

    if (!this.hasValidTrackChangeContract()) {
      const invalidContractResult = this.resultFactory.buildErrorResult(
        "Contrato invalido de track-change: anchor y suggestedText son obligatorios.",
        pendingBefore,
        "observe-before",
      );
      return this.toOutcome(invalidContractResult, pendingBefore);
    }

    await this.observabilityReporter.emitPhase("observe-before", "started", {
      suggestionType: this.suggestion.type,
    });

    const observation = await this.observer.observeResolutionCandidates(
      context,
      candidates,
      selectedCc,
    );

    if (observation.observationStatus !== "confirmed-pending") {
      const failureStatus =
        observation.observationStatus === "identity-lost" ||
        observation.observationStatus === "ambiguous-location" ||
        observation.observationStatus === "mixed-group" ||
        observation.observationStatus === "unobservable"
          ? observation.observationStatus
          : "unobservable";
      await this.observabilityReporter.emitPhase(
        "observe-before",
        failureStatus === "unobservable" ? "warning" : "failed",
        this.observabilityReporter.mergeMetadata(
          { reason: failureStatus },
          observation.debugMetadata,
        ),
      );

      const result = await this.resultFactory.buildObservationFailureResult(
        context,
        failureStatus,
        pendingBefore,
      );
      return this.toOutcome(result, pendingBefore);
    }

    await this.observabilityReporter.emitPhase(
      "observe-before",
      "succeeded",
      this.observabilityReporter.mergeMetadata(
        { trackedChangesObserved: observation.trackedChanges.length },
        observation.debugMetadata,
      ),
    );

    const executionReport = await this.executeOperationalScope(
      context,
      observation,
    );
    if (executionReport.error) {
      const errorResult = this.resultFactory.buildErrorResult(
        executionReport.error,
        await this.stateInspector.inspect(context),
        "execute",
        executionReport,
      );
      return this.toOutcome(errorResult, pendingBefore);
    }

    const commentDeleted =
      await this.cleanup.deleteLocatedStylisticCommentAfterResolution(
        context,
        observation.selectedComment,
      );
    const pendingAfter =
      await this.stateInspector.inspectAfterResolution(context);

    await this.observabilityReporter.emitPhase("cleanup-comment", "succeeded", {
      commentDeleted,
    });
    await this.observabilityReporter.emitPhase("inspect-after", "succeeded", {
      pendingArtifacts: pendingAfter.pendingStylisticArtifacts,
    });

    return {
      status: this.resultFactory.toResolutionStatus(),
      trackedChangesAffected: executionReport.completed,
      commentDeleted,
      pendingBefore,
      pendingAfter,
      executionReport,
    };
  }

  /** Resolves comment-only suggestions without entering the tracked-change path. */
  private async resolveCommentOnly(
    context: Word.RequestContext,
    selectedCc: Word.ContentControl,
    pendingBefore: DocumentReviewState,
  ): Promise<CohesiveResolutionOutcome> {
    const colocatedComment = await this.locator.findColocatedStylisticComment(
      context,
      selectedCc,
    );
    const commentDeleted = await this.cleanup.deleteLocatedStylisticComment(
      context,
      colocatedComment,
    );
    const result = await this.commentOnlyResolver.resolve({
      context,
      cc: selectedCc,
      commentDeleted,
      pendingBefore,
    });

    return this.toOutcome(result, pendingBefore);
  }

  /** Executes one action against the wrapper-owned tracked-change collection. */
  private async executeOperationalScope(
    context: Word.RequestContext,
    observation: ResolutionObservation,
  ): Promise<ResolutionExecutionReport> {
    const collection = observation.trackedChangesCollection;
    if (!collection) {
      return {
        attempted: 0,
        completed: 0,
        remaining: 0,
        error:
          "La sugerencia no expuso una colección operacional ejecutable dentro del wrapper.",
      };
    }

    await this.observabilityReporter.emitPhase("execute", "started", {
      trackedChangesAttempted: observation.trackedChanges.length,
    });

    try {
      if (this.action === "accept") {
        collection.acceptAll();
      } else {
        collection.rejectAll();
      }
      await context.sync();
    } catch (error) {
      const serialized = this.errorSerializer.serialize(error);
      await this.observabilityReporter.emitPhase("execute", "failed", {
        attempted: observation.trackedChanges.length,
        completed: 0,
        remaining: observation.trackedChanges.length,
        error: serialized.message,
      });

      return {
        attempted: observation.trackedChanges.length,
        completed: 0,
        remaining: observation.trackedChanges.length,
        failureIndex: 0,
        error: serialized.message,
      };
    }

    const report: ResolutionExecutionReport = {
      attempted: observation.trackedChanges.length,
      completed: observation.trackedChanges.length,
      remaining: 0,
    };
    await this.observabilityReporter.emitPhase("execute", "succeeded", {
      attempted: report.attempted,
      completed: report.completed,
      remaining: report.remaining,
    });
    return report;
  }

  /** Returns true only when the current suggestion satisfies the replace contract. */
  private hasValidTrackChangeContract(): boolean {
    return (
      this.suggestion.type === "track-change" &&
      this.suggestion.anchor.trim().length > 0 &&
      (this.suggestion.suggestedText?.trim().length ?? 0) > 0
    );
  }

  /** Normalizes a public result into the private cohesive outcome shape. */
  private toOutcome(
    result: SuggestionActionResult,
    pendingBefore: DocumentReviewState,
  ): CohesiveResolutionOutcome {
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
