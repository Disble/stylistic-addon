import type {
  Suggestion,
  SuggestionObservationStatus,
} from "../../../domain/suggestion/Suggestion.types";
import {
  STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../../infrastructure/config";
import {
  isValidOperationalReplaceIdentity,
  parseReplaceIdentityTitle,
} from "../ReplaceIdentityParser";
import type { LocatedSuggestionArtifactResult } from "./SuggestionArtifactLocator.types";

/**
 * Locates persisted Stylistic Word artifacts without mutating or selecting.
 *
 * This service is shared by navigation and resolution so both workflows agree on
 * which persisted Word artifact is safe to trust. It does not observe tracked
 * changes, accept/reject, delete comments, or select ranges; callers decide what
 * to do after a safe artifact is found.
 */
export class SuggestionArtifactLocator {
  /**
   * Finds the strict operational wrapper used by track-change suggestions.
   *
   * A candidate is safe only when its `compound-v2`/operational-wrapper metadata
   * validates against the current suggestion. Duplicate valid candidates are
   * ambiguous; malformed Stylistic metadata becomes `identity-lost`.
   */
  async locateOperationalWrapper(
    context: Word.RequestContext,
    suggestion: Suggestion
  ): Promise<LocatedSuggestionArtifactResult> {
    const { candidates } = await this.locateByTag(
      context,
      `${STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX}${suggestion.id}`
    );
    const validCandidates = candidates.filter((cc) =>
      isValidOperationalReplaceIdentity(parseReplaceIdentityTitle(cc.title), suggestion)
    );

    return {
      candidates,
      selectedCc: validCandidates.length === 1 ? validCandidates[0] : null,
      locateStatus: this.resolveOperationalWrapperStatus(candidates, validCandidates),
    };
  }

  /**
   * Finds the canonical anchor Content Control used by comment-only suggestions.
   *
   * Comment-only suggestions do not have operational-wrapper metadata, so exact
   * tag uniqueness is the safety contract.
   */
  async locateCommentOnlyArtifact(
    context: Word.RequestContext,
    suggestion: Suggestion
  ): Promise<LocatedSuggestionArtifactResult> {
    return this.locateByTag(context, `${STYLISTIC_TAG_PREFIX}comment-only:${suggestion.id}`);
  }

  /** Loads all Content Controls matching one exact tag. */
  private async locateByTag(
    context: Word.RequestContext,
    tag: string
  ): Promise<LocatedSuggestionArtifactResult> {
    const result = context.document.contentControls.getByTag(tag);
    result.load("items");
    result.load("items/tag,items/title");
    await context.sync();

    const candidates = result.items;
    return {
      candidates,
      selectedCc: candidates.length === 1 ? candidates[0] : null,
      locateStatus: this.resolveStrictTagStatus(candidates),
    };
  }

  /** Classifies operational-wrapper lookup after identity validation. */
  private resolveOperationalWrapperStatus(
    candidates: Word.ContentControl[],
    validCandidates: Word.ContentControl[]
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
      (cc.title ?? "").startsWith("stylistic-meta-v2:")
    );
    return hasMalformedOperationalMetadata ? "identity-lost" : "ambiguous-location";
  }

  /** Classifies exact-tag lookup without ranking fallback. */
  private resolveStrictTagStatus(
    candidates: Word.ContentControl[]
  ): SuggestionObservationStatus | "cc-not-found" {
    if (candidates.length === 0) {
      return "cc-not-found";
    }

    return candidates.length === 1 ? "confirmed-pending" : "ambiguous-location";
  }
}
