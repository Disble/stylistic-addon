import { DocumentReviewStateMachine } from "../../../domain/review/DocumentReviewStateMachine";
import type { DocumentReviewState } from "../../../domain/review/DocumentReviewStateMachine.types";
import type {
  ResolutionExecutionReport,
  ResolutionPhase,
  SuggestionActionResult,
} from "../../../domain/suggestion/SuggestionResolutionWorkflow.types";
import type { ResolutionStatus } from "./ResolutionContext";

/** Shapes stable taskpane-facing results for resolution workflows. */
export class ResolveSuggestionResultFactory {
  constructor(
    private readonly action: "accept" | "reject",
    private readonly stateInspector: {
      deriveDocumentState(
        reviewState: DocumentReviewState,
      ): import("../../../domain/review/DocumentReviewStateMachine").DocumentReviewUiState;
      inspect(context: Word.RequestContext): Promise<DocumentReviewState>;
    },
  ) {}

  /** Maps a resolution action to its terminal success status. */
  toResolutionStatus(): ResolutionStatus {
    return this.action === "accept"
      ? ("accepted" as const)
      : ("rejected" as const);
  }

  /** Builds a document-aware resolution result. */
  buildResolutionResult(
    status: SuggestionActionResult["status"],
    trackedChangesAffected: number,
    commentDeleted: boolean,
    pendingBefore: DocumentReviewState,
    pendingAfter: DocumentReviewState,
    error?: string,
    executionReport?: ResolutionExecutionReport,
  ): SuggestionActionResult {
    const transition = DocumentReviewStateMachine.evaluateTransition(
      pendingBefore,
      pendingAfter,
    );

    return {
      status,
      trackedChangesAffected,
      commentDeleted,
      pendingAfter,
      documentState: transition.to,
      ...(error ? { error } : {}),
      ...(executionReport ? { executionReport } : {}),
    };
  }

  /** Builds an observation failure result for fail-closed host evidence. */
  async buildObservationFailureResult(
    context: Word.RequestContext,
    status:
      | "identity-lost"
      | "unobservable"
      | "ambiguous-location"
      | "mixed-group",
    pendingBefore: DocumentReviewState,
  ): Promise<SuggestionActionResult> {
    const pendingAfter = await this.stateInspector.inspect(context);
    const error = this.buildObservationFailureMessage(status);

    return this.buildResolutionResult(
      status,
      0,
      false,
      pendingBefore,
      pendingAfter,
      error,
    );
  }

  /** Converts explicit observation failure status into user-facing Spanish copy. */
  private buildObservationFailureMessage(
    status:
      | "identity-lost"
      | "unobservable"
      | "ambiguous-location"
      | "mixed-group",
  ): string {
    if (status === "identity-lost") {
      return "La metadata operational-wrapper de la sugerencia está incompleta o corrupta.";
    }

    if (status === "ambiguous-location") {
      return "La ubicación de la sugerencia es ambigua; se abortó antes de modificar el documento.";
    }

    if (status === "mixed-group") {
      return "El grupo contiguo requiere resolución grupal coherente; se abortó antes de modificar el documento.";
    }

    return "Word no expuso suficiente evidencia operacional para confirmar la resolución.";
  }

  /** Builds a stable outer-catch error result. */
  buildErrorResult(
    error: string,
    pendingAfter: DocumentReviewState,
    errorPhase?: ResolutionPhase,
    executionReport?: ResolutionExecutionReport,
  ): SuggestionActionResult {
    return {
      status: "error",
      trackedChangesAffected: executionReport?.completed ?? 0,
      commentDeleted: false,
      pendingAfter,
      documentState: this.stateInspector.deriveDocumentState(pendingAfter),
      error,
      ...(errorPhase ? { errorPhase } : {}),
      ...(executionReport ? { executionReport } : {}),
    };
  }
}
