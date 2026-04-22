import type {
  DocumentReviewState,
  ResolutionExecutionReport,
  SuggestionActionResult,
  SuggestionResolutionWarning,
} from "../../../domain/types";
import type { ResolveSuggestionResultFactory } from "./ResolveSuggestionResultFactory";

/**
 * Reconciles late host failures against already-completed semantic resolution.
 *
 * Word can mutate tracked changes successfully and only fail later during
 * cleanup or post-resolution reads. This collaborator preserves semantic truth:
 * if execution completed with no remaining tracked changes, the taskpane must
 * receive a terminal result plus warnings instead of a retryable lie.
 */
export class SuggestionResolutionResolver {
  constructor(
    private readonly suggestionId: string,
    private readonly resultFactory: ResolveSuggestionResultFactory,
  ) {}

  /**
   * Converts a late failure into a terminal semantic result when execution had
   * already resolved every tracked change.
   */
  reconcileLateFailure(
    pendingAfter: DocumentReviewState,
    executionReport: ResolutionExecutionReport | undefined,
    message: string,
    inheritedWarnings: SuggestionResolutionWarning[] = [],
  ): SuggestionActionResult | null {
    if (!executionReport || executionReport.remaining > 0) {
      return null;
    }

    console.warn(
      `⚠️ [SuggestionResolutionResolver] "${this.suggestionId}": preserving semantic ${this.resultFactory.toResolutionStatus()} after late host failure: ${message}`,
    );

    return this.resultFactory.buildResolutionResult(
      this.resultFactory.toResolutionStatus(),
      executionReport.completed,
      false,
      pendingAfter,
      pendingAfter,
      undefined,
      [
        ...inheritedWarnings,
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
      executionReport,
    );
  }
}
