/**
 * Application constants for the Stylistic add-in.
 *
 * All configurable values live here as named exports. No UI-facing settings —
 * these are internal defaults that control backend communication, text chunking,
 * Word API batching, and retry behavior.
 *
 * @module config
 */

import type { Profile } from "../domain/types";

// ---------------------------------------------------------------------------
// Mastra Backend
// ---------------------------------------------------------------------------

/** Base URL of the Mastra server. */
export const MASTRA_BASE_URL = "http://localhost:4111";

/** Identifier of the editorial workflow registered in Mastra. */
export const WORKFLOW_ID = "stylistic-workflow";

/** Identifier of the feedback workflow registered in Mastra. */
export const FEEDBACK_WORKFLOW_ID = "feedback-workflow";

// ---------------------------------------------------------------------------
// Text Chunking
// ---------------------------------------------------------------------------

/**
 * Maximum character count per chunk sent to the workflow.
 * Paragraphs are accumulated until this limit is reached.
 * A single paragraph exceeding this size is sent as-is.
 */
export const DEFAULT_MAX_CHUNK_SIZE = 100_000;

// ---------------------------------------------------------------------------
// Word API Batching
// ---------------------------------------------------------------------------

/**
 * Number of suggestions applied per `Word.run` batch.
 * Each batch is an independent commit — if batch N fails,
 * batches 1..(N-1) are already persisted in the document.
 */
export const WORD_API_BATCH_SIZE = 30;

// ---------------------------------------------------------------------------
// Retry Policy
// ---------------------------------------------------------------------------

/** Maximum number of retry attempts per chunk on transient failures. */
export const MAX_RETRIES = 3;

/** Base delay (ms) for exponential backoff between retries: base * 2^attempt. */
export const RETRY_BASE_DELAY_MS = 1_000;

/** Delay (ms) between round-robin polling passes while chunks are still running. */
export const POLL_INTERVAL_MS = 1_000;

// ---------------------------------------------------------------------------
// Word Artifact Identity
// ---------------------------------------------------------------------------

/** Prefix used by all Stylistic Content Control tags persisted in Word. */
export const STYLISTIC_TAG_PREFIX = "stylistic:";

/**
 * Prefix for versioned Content Control title metadata used by compound replace
 * identities. The raw JSON payload is appended after this marker.
 */
export const STYLISTIC_IDENTITY_TITLE_PREFIX = "stylistic-meta-v2:";

// ---------------------------------------------------------------------------
// Default Profiles
// ---------------------------------------------------------------------------

/**
 * Fallback analysis profiles shown in the dropdown when the backend
 * is unavailable or does not provide its own profile list.
 */
export const DEFAULT_PROFILES: Profile[] = [
  { id: "narrativa-literaria", label: "Literatura de ficción" },
  { id: "general", label: "General" },
  { id: "ensayo-academico", label: "Ensayo académico" },
  { id: "periodismo-cultural", label: "Periodismo cultural" },
];
