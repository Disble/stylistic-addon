import { DocumentReviewStateMachine } from "../../../domain/review/DocumentReviewStateMachine";
import type {
  DocumentReviewState,
  SuggestionActionResult,
} from "../../../domain/types";
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
    };
  }

  /** Builds an observation failure result for identity-lost or unobservable evidence. */
  async buildObservationFailureResult(
    context: Word.RequestContext,
    status: "identity-lost" | "unobservable",
    pendingBefore: DocumentReviewState,
  ): Promise<SuggestionActionResult> {
    const pendingAfter = await this.stateInspector.inspect(context);
    const error =
      status === "identity-lost"
        ? "La metadata compound-v2 de la sugerencia está incompleta o corrupta."
        : "Word no expuso suficientes tracked changes para confirmar la resolución.";

    return this.buildResolutionResult(
      status,
      0,
      false,
      pendingBefore,
      pendingAfter,
      error,
    );
  }

  /** Builds a stable outer-catch error result. */
  buildErrorResult(
    error: string,
    pendingAfter: DocumentReviewState,
  ): SuggestionActionResult {
    return {
      status: "error",
      trackedChangesAffected: 0,
      commentDeleted: false,
      pendingAfter,
      documentState: this.stateInspector.deriveDocumentState(pendingAfter),
      error,
    };
  }
}
