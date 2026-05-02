import type {
  ReplaceSuggestionIdentity,
  Suggestion,
  WordArtifactRef,
} from "../../../domain/suggestion/Suggestion.types";
import {
  STYLISTIC_IDENTITY_TITLE_PREFIX,
  STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../../infrastructure/config";
import { TrackChangeSubtypeResolver } from "./TrackChangeSubtypeResolver";

/**
 * Builds persisted identity artifacts used by `ApplySuggestionCommand`.
 */
export class ApplySuggestionIdentityBuilder {
  private readonly subtypeResolver = new TrackChangeSubtypeResolver();

  /** Returns the canonical Stylistic tag for one suggestion. */
  buildSuggestionTag(suggestion: Suggestion): string {
    return `${STYLISTIC_TAG_PREFIX}${suggestion.type}:${suggestion.id}`;
  }

  /** Returns the external operational-wrapper tag for one replace suggestion. */
  buildOperationalWrapperTag(suggestion: Suggestion): string {
    return `${STYLISTIC_OPERATIONAL_WRAPPER_TAG_PREFIX}${suggestion.id}`;
  }

  /** Creates one Word artifact reference owned by the apply adapter. */
  createArtifactRef(
    kind: WordArtifactRef["kind"],
    role: WordArtifactRef["role"],
    value: string,
  ): WordArtifactRef {
    return { kind, role, value };
  }

  /** Serializes versioned replace identity metadata into the Content Control title. */
  serializeReplaceIdentity(identity: ReplaceSuggestionIdentity): string {
    return `${STYLISTIC_IDENTITY_TITLE_PREFIX}${JSON.stringify(identity)}`;
  }

  /**
   * Builds strict operational-wrapper metadata for replace suggestions.
   *
   * The inserted-side Content Control remains an operational reference, not the
   * whole domain identity. Deleted/original-side and anchor references are stored
   * explicitly so later observation can distinguish legacy vs v2 behavior.
   */
  buildReplaceIdentity(suggestion: Suggestion): ReplaceSuggestionIdentity {
    const subtypeResolution = this.subtypeResolver.resolve(suggestion);
    const subtype =
      subtypeResolution.subtype === "insert"
        ? "replace"
        : subtypeResolution.subtype;
    const baseIdentity = {
      suggestionId: suggestion.id,
      version: "operational-wrapper-v1" as const,
      anchorRef: this.createArtifactRef(
        "anchor",
        "operational-anchor",
        suggestion.context,
      ),
      groupId: suggestion.id,
      groupIndex: 0,
      groupSize: 1,
    };

    if (subtype === "delete-only") {
      return {
        ...baseIdentity,
        trackChangeSubtype: "delete-only",
        deletedSideRef: this.createArtifactRef(
          "anchor",
          "delete-side",
          suggestion.anchor,
        ),
      };
    }

    if (subtype === "formatting") {
      return {
        ...baseIdentity,
        trackChangeSubtype: "formatting",
        formatSideRef: this.createArtifactRef(
          "content-control",
          "format-side",
          this.buildSuggestionTag(suggestion),
        ),
      };
    }

    return {
      version: "operational-wrapper-v1",
      suggestionId: suggestion.id,
      insertedSideRef: this.createArtifactRef(
        "content-control",
        "inserted-side",
        this.buildSuggestionTag(suggestion),
      ),
      deletedSideRef: this.createArtifactRef(
        "anchor",
        "deleted-side",
        suggestion.anchor,
      ),
      anchorRef: this.createArtifactRef(
        "anchor",
        "operational-anchor",
        suggestion.context,
      ),
      groupId: suggestion.id,
      groupIndex: 0,
      groupSize: 1,
    };
  }

  /** Chooses the persisted Content Control title payload for one suggestion. */
  buildContentControlTitle(suggestion: Suggestion): string {
    return suggestion.anchor;
  }
}
