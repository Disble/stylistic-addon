import type { Suggestion } from "../../../domain/types";
import { OVERLAPPING_RELATIONS } from "../cleanup/CommentCleanup";
import {
  isValidCompoundReplaceIdentity,
  parseReplaceIdentityTitle,
  scoreCompoundReplaceIdentityMatch,
} from "../ReplaceIdentityParser";
import { isStylisticComment } from "../StylisticCommentBuilder";
import type {
  ColocatedCommentContext,
  LocatedSuggestionArtifacts,
} from "./ResolutionContext";

type ResolutionCandidateDebugEntry = {
  candidateIndex: number;
  wasSelected: boolean;
  selectionReason: "valid-compound-v2" | "not-selected";
  tag: string;
  titleKind: "compound-v2" | "invalid-or-missing";
  rawTitle: string;
  score: number;
  validCompoundV2: boolean;
  identityVersion: string | null;
  identitySuggestionId: string | null;
  insertedSideValue: string | null;
  deletedSideValue: string | null;
  anchorValue: string | null;
};

/** Finds the right Word artifacts for one resolution workflow. */
export class SuggestionLocator {
  constructor(private readonly suggestion: Suggestion) {}

  /** Builds one structured debug entry so duplicate CC candidates can be compared. */
  private buildCandidateDebugEntry(
    cc: Word.ContentControl,
    candidateIndex: number,
    selectedCc: Word.ContentControl | null,
  ): ResolutionCandidateDebugEntry {
    const identity = parseReplaceIdentityTitle(cc.title);
    const validCompoundV2 = isValidCompoundReplaceIdentity(
      identity,
      this.suggestion,
    );
    const selectedBecauseCompoundV2 = selectedCc === cc && validCompoundV2;
    let selectionReason: ResolutionCandidateDebugEntry["selectionReason"] =
      "not-selected";
    if (selectedBecauseCompoundV2) {
      selectionReason = "valid-compound-v2";
    }

    return {
      candidateIndex,
      wasSelected: selectedCc === cc,
      selectionReason,
      tag: cc.tag,
      titleKind:
        identity?.version === "compound-v2"
          ? "compound-v2"
          : "invalid-or-missing",
      rawTitle: cc.title ?? "",
      score: scoreCompoundReplaceIdentityMatch(identity, this.suggestion),
      validCompoundV2,
      identityVersion: identity?.version ?? null,
      identitySuggestionId: identity?.suggestionId ?? null,
      insertedSideValue: identity?.insertedSideRef?.value ?? null,
      deletedSideValue: identity?.deletedSideRef?.value ?? null,
      anchorValue: identity?.anchorRef?.value ?? null,
    };
  }

  /** Logs all candidate details needed to tell a fresh CC from stale duplicates. */
  private logResolutionCandidateDiagnostics(
    rankedCandidates: Word.ContentControl[],
    selectedCc: Word.ContentControl | null,
  ): void {
    const diagnostics = rankedCandidates.map((cc, index) =>
      this.buildCandidateDebugEntry(cc, index, selectedCc),
    );

    console.log(
      `🧾 [SuggestionLocator] candidate diagnostics for suggestionId="${this.suggestion.id}"`,
      diagnostics,
    );

    const indistinguishableCandidates = diagnostics.filter((candidate) => {
      const comparableKey = JSON.stringify({
        tag: candidate.tag,
        rawTitle: candidate.rawTitle,
        score: candidate.score,
        validCompoundV2: candidate.validCompoundV2,
      });

      return (
        diagnostics.filter((otherCandidate) => {
          const otherComparableKey = JSON.stringify({
            tag: otherCandidate.tag,
            rawTitle: otherCandidate.rawTitle,
            score: otherCandidate.score,
            validCompoundV2: otherCandidate.validCompoundV2,
          });

          return comparableKey === otherComparableKey;
        }).length > 1
      );
    });

    if (indistinguishableCandidates.length > 1) {
      console.warn(
        `⚠️ [SuggestionLocator] indistinguishable duplicate candidates for suggestionId="${this.suggestion.id}"`,
        indistinguishableCandidates,
      );
    }
  }

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
    this.logResolutionCandidateDiagnostics(rankedCandidates, selectedCc);

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

    return v2Candidate ?? null;
  }

  /** Orders CC candidates so valid compound-v2 artifacts are tried first. */
  rankResolutionContentControls(
    ccs: Word.ContentControl[],
  ): Word.ContentControl[] {
    return [...ccs].sort((left, right) => {
      const leftScore = scoreCompoundReplaceIdentityMatch(
        parseReplaceIdentityTitle(left.title),
        this.suggestion,
      );
      const rightScore = scoreCompoundReplaceIdentityMatch(
        parseReplaceIdentityTitle(right.title),
        this.suggestion,
      );

      return rightScore - leftScore;
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
