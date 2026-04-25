import type {
  Suggestion,
  SuggestionObservationStatus,
} from "../../../domain/types";
import { STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX } from "../../../infrastructure/config";
import { OVERLAPPING_RELATIONS } from "../cleanup/CommentCleanup";
import {
  isValidOperationalReplaceIdentity,
  parseReplaceIdentityTitle,
} from "../ReplaceIdentityParser";
import { isStylisticComment } from "../StylisticCommentBuilder";
import type {
  ColocatedCommentContext,
  LocatedSuggestionArtifacts,
} from "./ResolutionContext";

/** Finds the exact Word artifacts for one operational-wrapper resolution workflow. */
export class SuggestionLocator {
  constructor(private readonly suggestion: Suggestion) {}

  /** Finds a single strict operational wrapper and rejects ambiguous duplicates before mutation. */
  async locateResolutionArtifacts(
    context: Word.RequestContext,
  ): Promise<LocatedSuggestionArtifacts> {
    const result = context.document.contentControls.getByTag(
      `${STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX}${this.suggestion.id}`,
    );
    result.load("items/tag,items/title");
    await context.sync();

    const candidates = result.items;
    const validCandidates = candidates.filter((cc) =>
      isValidOperationalReplaceIdentity(
        parseReplaceIdentityTitle(cc.title),
        this.suggestion,
      ),
    );
    const locateStatus = this.resolveLocateStatus(candidates, validCandidates);
    const selectedCc = validCandidates.length === 1 ? validCandidates[0] : null;

    console.log(
      `🎯 [SuggestionLocator] strict operational lookup suggestionId="${this.suggestion.id}" candidates=${candidates.length} valid=${validCandidates.length} status=${locateStatus}`,
    );

    return { candidates, selectedCc, locateStatus };
  }

  /** Classifies lookup results without ranking or compatibility fallback. */
  private resolveLocateStatus(
    candidates: Word.ContentControl[],
    validCandidates: Word.ContentControl[],
  ): SuggestionObservationStatus | "cc-not-found" {
    if (candidates.length === 0) {
      return "cc-not-found";
    }

    if (validCandidates.length === 1) {
      return "confirmed-pending";
    }

    if (validCandidates.length > 1) {
      return "ambiguous-location";
    }

    const hasMalformedOperationalMetadata = candidates.some((cc) =>
      (cc.title ?? "").startsWith("stylistic-meta-v2:"),
    );
    return hasMalformedOperationalMetadata
      ? "identity-lost"
      : "ambiguous-location";
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

    for (const comment of stylisticComments) {
      const commentRange = comment.getRange();
      const locationResult = commentRange.compareLocationWith(ccRange);
      await context.sync();

      if (OVERLAPPING_RELATIONS.includes(locationResult.value as string)) {
        return { comment, range: commentRange };
      }
    }

    return null;
  }
}
