import type { Suggestion } from "../../../domain/suggestion/Suggestion.types";
import { isOverlappingRelation } from "../cleanup/CommentCleanup";
import { SuggestionArtifactLocator } from "../location/SuggestionArtifactLocator";
import { isStylisticComment } from "../StylisticCommentBuilder";
import type { ColocatedCommentContext, LocatedSuggestionArtifacts } from "./ResolutionContext";

/** Finds the exact Word artifacts for one operational-wrapper resolution workflow. */
export class SuggestionLocator {
  private readonly artifactLocator = new SuggestionArtifactLocator();

  constructor(private readonly suggestion: Suggestion) {}

  /** Finds the unique comment-only anchor Content Control using its canonical tag. */
  async locateCommentOnlyArtifacts(
    context: Word.RequestContext
  ): Promise<LocatedSuggestionArtifacts> {
    const result = await this.artifactLocator.locateCommentOnlyArtifact(context, this.suggestion);

    console.log(
      `🎯 [SuggestionLocator] comment-only lookup suggestionId="${this.suggestion.id}" candidates=${result.candidates.length} status=${result.locateStatus}`
    );

    return result;
  }

  /** Finds a single strict operational wrapper and rejects ambiguous duplicates before mutation. */
  async locateResolutionArtifacts(
    context: Word.RequestContext
  ): Promise<LocatedSuggestionArtifacts> {
    const result = await this.artifactLocator.locateOperationalWrapper(context, this.suggestion);

    console.log(
      `🎯 [SuggestionLocator] strict operational lookup suggestionId="${this.suggestion.id}" candidates=${result.candidates.length} selected=${result.selectedCc ? 1 : 0} status=${result.locateStatus}`
    );

    return result;
  }

  /** Finds the Stylistic comment colocated with the suggestion CC range. */
  async findColocatedStylisticComment(
    context: Word.RequestContext,
    cc: Word.ContentControl
  ): Promise<ColocatedCommentContext | null> {
    const comments = context.document.body.getComments();
    comments.load({ select: "authorName,content" });
    await context.sync();

    const stylisticComments = comments.items.filter(isStylisticComment);
    const ccRange = cc.getRange();

    for (const comment of stylisticComments) {
      const commentRange = comment.getRange();
      const locationResult = commentRange.compareLocationWith(ccRange);
      await context.sync();

      if (isOverlappingRelation(locationResult.value as string)) {
        return { comment, range: commentRange };
      }
    }

    return null;
  }
}
