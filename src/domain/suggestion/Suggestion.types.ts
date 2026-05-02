/**
 * Core suggestion contracts shared by the domain, adapters, and taskpane.
 */

/** Supported editorial severity levels across suggestions and feedback. */
export type SuggestionSeverity = "high" | "medium" | "low";

/** Supported suggestion materialization modes in Word. */
export type SuggestionType = "track-change" | "comment-only";

/**
 * Frontend-owned position hint derived from a live apply snapshot.
 */
export interface SuggestionBatchPositionHint {
  /** Start offset relative to the coordinate source that produced this hint. */
  start: number;

  /** End offset relative to the coordinate source that produced this hint. */
  end: number;

  /** Snapshot version whose coordinates this hint is comparable against. */
  snapshotVersion: number;

  /** Optional paragraph/local-container identity for localized recovery. */
  paragraphId?: string;

  /** Identifies the coordinate/ranking source that produced this hint. */
  source: "snapshot" | "localized-reread";

  /** Indicates the current hint must be revalidated with a localized reread. */
  requiresLocalReread?: boolean;
}

/**
 * A single editorial suggestion, either received from the Mastra workflow
 * or prepared for insertion into the Word document.
 */
export interface Suggestion {
  /** Unique identifier assigned by the frontend (e.g., "chunk0-3"). */
  id: string;

  /** Paragraph-level context used to locate the suggestion in the document. */
  context: string;

  /** Exact substring within `context` targeted by the suggestion. */
  anchor: string;

  /**
   * Transport text for tracked-change suggestions.
   *
   * - replacement text for normal replace suggestions,
   * - empty string for delete-only suggestions,
   * - markdown `*anchor*` / `**anchor**` for typography suggestions decoded by
   *   the Word adapter into native italic/bold formatting.
   */
  suggestedText?: string;

  /** Human-readable reason for the suggestion, shown in the results panel. */
  justification: string;

  /** Editorial category label (e.g., "Redundancia", "Muletilla"). */
  category: string;

  /** How critical the suggestion is. */
  severity: SuggestionSeverity;

  /** Determines how the suggestion is applied to the document. */
  type: SuggestionType;

  /** Optional live position hint used by batch apply orchestration. */
  positionHint?: SuggestionBatchPositionHint;
}

/** User-visible result of trying to move the Word selection to a suggestion. */
export type SuggestionNavigationResult =
  | { status: "navigated" }
  | {
      status: "not-found";
      reason:
        | "artifact-not-found"
        | "context-not-found"
        | "anchor-not-found"
        | "plain-text-not-found";
    }
  | {
      status: "ambiguous";
      reason: "multiple-artifacts" | "identity-lost" | "mixed-group";
    }
  | { status: "failed"; reason: "word-error" };

/** Observation confidence for a suggestion materialized in Word. */
export type SuggestionObservationStatus =
  | "confirmed-pending"
  | "confirmed-resolved"
  | "unobservable"
  | "identity-lost"
  | "ambiguous-location"
  | "mixed-group";

/** Fail-closed reasons emitted before a replace resolution mutates Word. */
export type ResolutionAbortReason = "ambiguous-location" | "mixed-group";

/**
 * A Word-host reference that helps re-locate one side of a review suggestion.
 */
export interface WordArtifactRef {
  /** Kind of Word artifact being referenced. */
  kind: "content-control" | "tracked-change" | "comment" | "anchor";

  /** Semantic role this artifact plays inside the suggestion identity. */
  role:
    | "inserted-side"
    | "deleted-side"
    | "delete-side"
    | "format-side"
    | "operational-anchor";

  /** Opaque adapter-owned value used to relocate the artifact in Word. */
  value: string;
}

/** Supported native Track Changes operational subtypes. */
export type TrackChangeSuggestionSubtype =
  | "replace"
  | "delete-only"
  | "formatting";

/** Versioned operational-wrapper identity for native Track Changes suggestions. */
export interface ReplaceSuggestionIdentity {
  /** Stable frontend/domain suggestion identifier. */
  suggestionId: string;

  /** Serialized identity version for strict operational-wrapper resolution. */
  version: "operational-wrapper-v1";

  /**
   * Native Track Changes subtype represented by this wrapper.
   *
   * Omitted legacy payloads are interpreted as `replace` to preserve already
   * persisted operational-wrapper identities.
   */
  trackChangeSubtype?: TrackChangeSuggestionSubtype;

  /** Primary inserted-side Word reference. */
  insertedSideRef?: WordArtifactRef;

  /** Primary formatting-side Word reference for formatting-only suggestions. */
  formatSideRef?: WordArtifactRef;

  /** Optional deleted/original-side Word reference. */
  deletedSideRef?: WordArtifactRef;

  /** Optional operational anchor for fallback re-location. */
  anchorRef?: WordArtifactRef;

  /** Explicit contiguous wrapper group identity. Defaults to the suggestion id. */
  groupId: string;

  /** Position of this wrapper inside its explicit contiguous group. */
  groupIndex: number;

  /** Number of wrappers expected in the explicit contiguous group. */
  groupSize: number;
}

/** Visual state of a suggestion card in the taskpane after user action. */
export type SuggestionState =
  | "pending"
  | "resolving"
  | "accepted"
  | "rejected"
  | "unobservable"
  | "identity-lost"
  | "ambiguous-location"
  | "mixed-group"
  | "error";
