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
 *
 * IMPORTANT:
 * For replace resolution we now prefer explicit, deterministic references over
 * legacy host-drift tolerance. The inserted-side tag must match the suggestion,
 * and both deleted-side plus operational-anchor refs must exist structurally so
 * the adapter can re-localize the replace without fuzzy fallback heuristics.
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
    identity.deletedSideRef.value.trim().length > 0 &&
    identity.anchorRef?.role === "operational-anchor" &&
    identity.anchorRef.value.trim().length > 0
  );
}

/** Returns the persisted deleted-side locator when the compound-v2 identity is valid. */
export function getDeletedSideLocator(
  identity: ReplaceSuggestionIdentity | null,
  suggestion: Suggestion,
): string | null {
  if (!isValidCompoundReplaceIdentity(identity, suggestion)) {
    return null;
  }

  return identity.deletedSideRef.value.trim() || null;
}

/** Returns the persisted operational-anchor locator when the compound-v2 identity is valid. */
export function getOperationalAnchorLocator(
  identity: ReplaceSuggestionIdentity | null,
  suggestion: Suggestion,
): string | null {
  if (!isValidCompoundReplaceIdentity(identity, suggestion)) {
    return null;
  }

  return identity.anchorRef.value.trim() || null;
}

/**
 * Scores how specifically one structurally valid compound-v2 identity matches
 * the current suggestion payload.
 *
 * The inserted-side CC tag remains the primary truth. Exact deleted/anchor text
 * matches are treated as affinity signals so the runtime can prefer the most
 * up-to-date artifact when duplicate CCs exist, without rejecting recoverable
 * host drift outright.
 */
export function scoreCompoundReplaceIdentityMatch(
  identity: ReplaceSuggestionIdentity | null,
  suggestion: Suggestion,
): number {
  if (!isValidCompoundReplaceIdentity(identity, suggestion)) {
    return -1;
  }

  let score = 1;

  if (identity.deletedSideRef?.value === suggestion.anchor) {
    score += 1;
  }

  if (identity.anchorRef?.value === suggestion.context) {
    score += 1;
  }

  return score;
}
