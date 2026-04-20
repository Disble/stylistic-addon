import type { ReplaceSuggestionIdentity, Suggestion } from "../../domain/types";
import {
  STYLISTIC_IDENTITY_TITLE_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../infrastructure/config";

/**
 * Parses persisted compound-v2 replace identity metadata from a Content Control
 * title payload.
 *
 * Returns `null` when the title does not carry a valid Stylistic identity
 * prefix or when the JSON payload is malformed.
 */
export function parseReplaceIdentityTitle(
  title: string | undefined,
): ReplaceSuggestionIdentity | null {
  const trimmed = title?.trim() ?? "";
  if (!trimmed.startsWith(STYLISTIC_IDENTITY_TITLE_PREFIX)) {
    return null;
  }

  const rawPayload = trimmed.slice(STYLISTIC_IDENTITY_TITLE_PREFIX.length);
  if (rawPayload.length === 0) {
    return null;
  }

  try {
    return JSON.parse(rawPayload) as ReplaceSuggestionIdentity;
  } catch {
    return null;
  }
}

/**
 * Validates the minimum compound-v2 replace identity contract required for
 * safe resolution and navigation.
 */
export function isValidCompoundReplaceIdentity(
  identity: ReplaceSuggestionIdentity | null,
  suggestion: Suggestion,
): identity is ReplaceSuggestionIdentity {
  if (identity?.version !== "compound-v2") {
    return false;
  }

  const expectedTag = `${STYLISTIC_TAG_PREFIX}${suggestion.type}:${suggestion.id}`;

  return (
    identity.suggestionId === suggestion.id &&
    identity.insertedSideRef?.kind === "content-control" &&
    identity.insertedSideRef.role === "inserted-side" &&
    identity.insertedSideRef.value === expectedTag &&
    identity.deletedSideRef?.role === "deleted-side" &&
    identity.deletedSideRef.value === suggestion.anchor &&
    identity.anchorRef?.role === "operational-anchor" &&
    identity.anchorRef.value === suggestion.context
  );
}
