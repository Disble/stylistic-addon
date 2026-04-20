import type { SuggestionActionResult } from "../../../domain/types";
import type { DocumentReviewStateInspector } from "./DocumentReviewStateInspector";
import type { CommentOnlyResolutionRequest } from "./ResolutionContext";
import type { ResolveSuggestionResultFactory } from "./ResolveSuggestionResultFactory";

/** Owns the comment-only branch once the anchor CC and comment are known. */
export class CommentOnlySuggestionResolver {
  constructor(
    private readonly suggestionId: string,
    private readonly resultFactory: ResolveSuggestionResultFactory,
    private readonly stateInspector: DocumentReviewStateInspector,
  ) {}

  /** Resolves a comment-only suggestion by deleting the anchor and returning status. */
  async resolve({
    context,
    cc,
    commentDeleted,
    pendingBefore,
  }: CommentOnlyResolutionRequest): Promise<SuggestionActionResult> {
    cc.delete(true);
    await context.sync();
    const pendingAfter = await this.stateInspector.inspect(context);

    console.log(
      `🗨️ [CommentOnlySuggestionResolver] "${this.suggestionId}": comment-only resolved, comentario eliminado: ${commentDeleted}`,
    );

    return this.resultFactory.buildResolutionResult(
      this.resultFactory.toResolutionStatus(),
      0,
      commentDeleted,
      pendingBefore,
      pendingAfter,
    );
  }
}
