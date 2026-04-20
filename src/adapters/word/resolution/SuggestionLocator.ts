import type { Suggestion } from "../../../domain/types";
import { OVERLAPPING_RELATIONS } from "../cleanup/CommentCleanup";
import {
  isValidCompoundReplaceIdentity,
  parseReplaceIdentityTitle,
} from "../ReplaceIdentityParser";
import { isStylisticComment } from "../StylisticCommentBuilder";
import type {
  ColocatedCommentContext,
  LocatedSuggestionArtifacts,
} from "./ResolutionContext";

/** Finds the right Word artifacts for one resolution workflow. */
export class SuggestionLocator {
  constructor(private readonly suggestion: Suggestion) {}

  /** Finds, ranks, and selects content-control candidates for one suggestion. */
  async locateResolutionArtifacts(
    context: Word.RequestContext,
  ): Promise<LocatedSuggestionArtifacts> {
    const result = context.document.contentControls.getByTag(
      `stylistic:${this.suggestion.type}:${this.suggestion.id}`,
    );
    result.load("items/tag,items/title");
    await context.sync();

    console.log(
      `🎯 [SuggestionLocator] getByTag returned ${result.items.length} CC candidate(s) for suggestionId="${this.suggestion.id}"`,
    );

    const rankedCandidates = this.rankResolutionContentControls(result.items);
    const selectedCc = this.selectResolutionContentControl(rankedCandidates);

    return { rankedCandidates, selectedCc };
  }

  /** Chooses the best CC candidate when multiple artifacts share the same tag. */
  selectResolutionContentControl(
    ccs: Word.ContentControl[],
  ): Word.ContentControl | null {
    if (ccs.length === 0) {
      return null;
    }

    const v2Candidate = ccs.find((cc) => {
      const identity = parseReplaceIdentityTitle(cc.title);
      return isValidCompoundReplaceIdentity(identity, this.suggestion);
    });

    return v2Candidate ?? ccs[0] ?? null;
  }

  /** Orders CC candidates so valid compound-v2 artifacts are tried first. */
  rankResolutionContentControls(
    ccs: Word.ContentControl[],
  ): Word.ContentControl[] {
    return [...ccs].sort((left, right) => {
      const leftValid = isValidCompoundReplaceIdentity(
        parseReplaceIdentityTitle(left.title),
        this.suggestion,
      );
      const rightValid = isValidCompoundReplaceIdentity(
        parseReplaceIdentityTitle(right.title),
        this.suggestion,
      );

      return Number(rightValid) - Number(leftValid);
    });
  }

  /** Finds the Stylistic comment colocated with the suggestion CC range. */
  async findColocatedStylisticComment(
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ): Promise<ColocatedCommentContext | null> {
    const comments = context.document.body.getComments();
    comments.load({ select: "authorName,content" });
    await context.sync();

    const stylisticComments = comments.items.filter(isStylisticComment);
    const ccRange = cc.getRange();

    console.log(
      `🔎 [SuggestionLocator] searching colocated comment for CC "${cc.tag}" among ${stylisticComments.length} Stylistic comments`,
    );

    for (const comment of stylisticComments) {
      const commentRange = comment.getRange();
      const locationResult = commentRange.compareLocationWith(ccRange);
      await context.sync();

      if (OVERLAPPING_RELATIONS.includes(locationResult.value as string)) {
        console.log(
          `🔎 [SuggestionLocator] found colocated comment for CC "${cc.tag}"`,
        );
        return { comment, range: commentRange };
      }
    }

    console.log(
      `🔎 [SuggestionLocator] no colocated comment found for CC "${cc.tag}"`,
    );
    return null;
  }
}
