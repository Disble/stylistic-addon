import type {
  ReplaceSuggestionIdentity,
  Suggestion,
} from "../../domain/suggestion/Suggestion.types";
import {
  STYLISTIC_IDENTITY_TITLE_PREFIX,
  STYLISTIC_TAG_PREFIX,
} from "../../infrastructure/config";

/**
 * Parses persisted operational-wrapper replace identity metadata from a Content
 * Control title payload.
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
 * Validates the minimum operational-wrapper replace identity contract required for
 * safe resolution and navigation.
 *
 * IMPORTANT:
 * For replace resolution we now prefer explicit, deterministic references over
 * legacy host-drift tolerance. The inserted-side tag must match the suggestion,
 * and both deleted-side plus operational-anchor refs must exist structurally so
 * the adapter can re-localize the replace without fuzzy fallback heuristics.
 */
export function isValidOperationalReplaceIdentity(
  identity: ReplaceSuggestionIdentity | null,
  suggestion: Suggestion,
): identity is ReplaceSuggestionIdentity {
  if (identity?.version !== "operational-wrapper-v1") {
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
    identity.anchorRef.value === suggestion.context &&
    identity.groupId.trim().length > 0 &&
    Number.isInteger(identity.groupIndex) &&
    identity.groupIndex >= 0 &&
    Number.isInteger(identity.groupSize) &&
    identity.groupSize >= 1 &&
    identity.groupIndex < identity.groupSize
  );
}

/** Validates versioned operational-wrapper metadata without binding it to one suggestion id. */
export function isStructurallyValidOperationalWrapperIdentity(
  identity: ReplaceSuggestionIdentity | null,
): identity is ReplaceSuggestionIdentity {
  return (
    identity?.version === "operational-wrapper-v1" &&
    identity.suggestionId.trim().length > 0 &&
    identity.insertedSideRef?.kind === "content-control" &&
    identity.insertedSideRef.role === "inserted-side" &&
    identity.insertedSideRef.value.trim().length > 0 &&
    identity.deletedSideRef?.role === "deleted-side" &&
    identity.deletedSideRef.value.trim().length > 0 &&
    identity.anchorRef?.role === "operational-anchor" &&
    identity.anchorRef.value.trim().length > 0 &&
    identity.groupId.trim().length > 0 &&
    Number.isInteger(identity.groupIndex) &&
    identity.groupIndex >= 0 &&
    Number.isInteger(identity.groupSize) &&
    identity.groupSize >= 1 &&
    identity.groupIndex < identity.groupSize
  );
}

/** Returns the persisted deleted-side locator when the operational-wrapper identity is valid. */
export function getDeletedSideLocator(
  identity: ReplaceSuggestionIdentity | null,
  suggestion: Suggestion,
): string | null {
  if (!isValidOperationalReplaceIdentity(identity, suggestion)) {
    return null;
  }

  return identity.deletedSideRef.value.trim() || null;
}

/** Returns the persisted operational-anchor locator when the operational-wrapper identity is valid. */
export function getOperationalAnchorLocator(
  identity: ReplaceSuggestionIdentity | null,
  suggestion: Suggestion,
): string | null {
  if (!isValidOperationalReplaceIdentity(identity, suggestion)) {
    return null;
  }

  return identity.anchorRef.value.trim() || null;
}
