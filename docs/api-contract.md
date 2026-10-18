# API Contract — Stylistic Hub API v1

This document records what the Stylistic Word add-in uses from Stylistic Hub
API v1. The hub owns the contract: its source of truth is the hub's
`docs/api.md` and the OpenAPI document it serves at `/api/v1/openapi.json`.
This page only describes the add-in's side of it.

## Connection

- **Base URL:** the task pane's own origin plus `/api/v1` (the hub serves the
  task pane, for example `https://localhost:3000/api/v1`).
- **Authentication:** there are no accounts. The hub injects its API token into
  the task pane as `<meta name="stylistic-hub-token" content="...">`; the add-in
  sends it as `Authorization: Bearer <token>` on every request
  (`src/infrastructure/hub/HubConnection.helpers.ts`).
- **No token:** when the task pane was not served by the hub, analysis refuses to
  start and tells the user to open the pane from the hub. There is no fallback
  backend.
- **Errors:** every refusal is `application/problem+json`
  `{ type, title, status, detail, code }`. The add-in shows `detail` (falling
  back to `title`) as the cause (`src/adapters/hub/HubProblemParser.helpers.ts`).

| Code                | Status   | Meaning for the add-in                       |
| ------------------- | -------- | -------------------------------------------- |
| `unauthorized`      | 401      | Missing or wrong hub token.                  |
| `validation_failed` | 400, 415 | The request body broke the contract.         |
| `not_found`         | 404      | Unknown analysis, suggestion or feedback id. |
| `no_model`          | 409      | The hub has no model configured to analyze.  |
| `payload_too_large` | 413      | The request exceeded the hub's size limit.   |
| `internal_error`    | 500      | The hub failed; `detail` carries the cause.  |

## Endpoints used

| Method and path                         | Adapter                               | Purpose                                    |
| --------------------------------------- | ------------------------------------- | ------------------------------------------ |
| `GET /status`                           | `HubAnalysisAdapter.getReadiness`     | Readiness check before every analysis.     |
| `POST /documents/{documentId}/analyses` | `HubAnalysisAdapter.startAnalysis`    | Starts one analysis for all chunks (202).  |
| `GET /analyses/{analysisId}`            | `HubAnalysisAdapter.getAnalysis`      | Polls the analysis until it stops running. |
| `POST /analyses/{analysisId}/cancel`    | `HubAnalysisAdapter.cancelAnalysis`   | Cancels the whole analysis.                |
| `POST /feedback`                        | `HubFeedbackAdapter.sendFeedback`     | Fire-and-forget accept/reject feedback.    |
| `GET /preferences`, `PUT /preferences`  | `HubUserCorrectionPreferencesAdapter` | Standing correction instructions.          |

## Analysis

### Readiness

`GET /status` returns `{ version, analysis: { ready, reason?, model? } }`. When
`analysis.ready` is false, the pipeline aborts before reading the document and
shows `analysis.reason` with a full-retry action. An unreachable hub counts as
not ready, with the transport error as the reason.

### Start

```json
POST /documents/11111111-1111-4111-8111-111111111111/analyses
{
  "genre": "general",
  "chunks": [
    { "id": "0", "text": "Básicamente, es completamente necesario..." },
    { "id": "1", "text": "Segundo fragmento..." }
  ]
}
```

- `documentId` is the stable document UUID the add-in persists in
  `Office.context.document.settings` (`stylistic.document.uuid.v1`).
- `genre` is the selected analysis profile: `narrativa-literaria`,
  `ensayo-academico`, `periodismo-cultural` or `general`.
- Chunk ids are the add-in's chunk indexes as strings (1–64 characters), so the
  hub's per-chunk results map back to the add-in's chunks. The add-in splits the
  document at paragraph boundaries into chunks of up to 100,000 characters; the
  hub may split them further internally.

### Poll and finish

`GET /analyses/{analysisId}` returns:

```typescript
interface HubAnalysis {
  id: string;
  documentId: string;
  status: "running" | "completed" | "partial" | "failed" | "canceled";
  chunks: Array<{
    id: string;
    status: "queued" | "running" | "completed" | "failed" | "canceled";
    suggestions: HubSuggestion[];
    dropped: number; // invalid or blank suggestions the hub discarded
    error?: string; // provider error for a failed chunk
  }>;
  profile: unknown;
  createdAt: string;
  finishedAt?: string;
}
```

The add-in polls every second while `status` is `running`, then:

- applies every suggestion from completed chunks, so a `partial` analysis still
  applies what succeeded;
- shows each failed chunk's `error` as the cause (prefixed with its fragment
  number when there is more than one chunk);
- treats `canceled` as a cancellation by the user;
- shows "No se encontraron sugerencias editoriales." when nothing failed and no
  suggestion came back.

A poll that fails on the add-in side keeps the analysis id and offers "retry the
query", which reads the same analysis again instead of starting a new one.
Analyses persist across hub restarts.

### Retries

`RetryAnalysisDecorator` retries `startAnalysis` and `getAnalysis` up to 3 times
with exponential backoff (1 s base) on transport errors, 5xx, 408 and 429. Any
other 4xx is a deterministic refusal and surfaces at once. Retrying
`startAnalysis` after a transport failure can create a duplicate analysis on the
hub if the first request did arrive; this is accepted.

## Suggestion Format

```typescript
interface HubSuggestion {
  /** 16 lowercase hex characters, issued by the hub, unique for good. */
  id: string;
  /** Paragraph-level context used to locate the suggestion in the document. */
  context: string;
  /** Exact substring within `context` targeted by the suggestion. */
  anchor: string;
  /**
   * Required for "track-change" and ignored for "comment-only":
   * - plain non-empty text replaces the anchor,
   * - an empty string deletes the anchor,
   * - exact markdown `*anchor*` / `**anchor**` applies italic / bold.
   */
  suggestedText?: string;
  justification: string;
  /** Category label, e.g. "Redundancia", "Muletilla". */
  category: string;
  severity: "high" | "medium" | "low";
  type: "track-change" | "comment-only";
}
```

The suggestion id goes into the Word content-control tag as
`stylistic:<type>:<id>` and is the key for feedback.

### Severity Guidelines

| Severity | When to use                                    | Example                         |
| -------- | ---------------------------------------------- | ------------------------------- |
| `high`   | Clear errors, redundancies, grammatical issues | "periodo de tiempo" → "periodo" |
| `medium` | Stylistic improvements, filler words           | "Básicamente, " → ""            |
| `low`    | Minor preferences, optional simplifications    | "utilizar" → "usar"             |

## Critical Constraint: context.includes(anchor)

The `anchor` in every suggestion **must** be an exact, character-for-character substring of the `context`. The frontend locates the context first, then searches for the anchor within that localized scope. If the anchor doesn't match exactly:

- The suggestion will silently fail (reported as "not found")
- No tracked change will be inserted for that suggestion
- Later cursor navigation may refuse to move because the add-in will not fall
  back to a global anchor search

This strictness is intentional. A global anchor search can select an unrelated
occurrence in a table of contents, heading, or repeated paragraph. The backend
contract must therefore provide enough context to make `context -> anchor`
localization safe.

**Do:**

- Return `"completamente necesario"` (exact match from input)
- Ensure `context.includes(anchor)` is true
- Include trailing punctuation/spaces if they're part of the replacement (e.g., `"Básicamente, "` to remove the comma and space)
- Return `suggestedText: ""` only when the intended change is to delete the
  exact anchor.
- Return exact typography markdown only when the inner text equals the anchor:
  `"*post mortem*"` for italic or `"**PRIME**"` for bold.

**Don't:**

- Return `"Completamente necesario"` (wrong capitalization)
- Return `"completamente  necesario"` (extra space)
- Return an anchor that is not contained in context
- Return markdown with different inner text than the anchor; the frontend treats
  that as a normal replacement, not a formatting instruction.
- Return literal asterisks when the desired result is typography. Markdown is a
  transport encoding for formatting, not visible document text.

## Genre Behavior

The `genre` field tells the hub which analysis style to apply:

| Genre                 | Description                   |
| --------------------- | ----------------------------- |
| `general`             | Everyday writing improvements |
| `narrativa-literaria` | Literary narrative style      |
| `ensayo-academico`    | Academic essay style          |
| `periodismo-cultural` | Cultural journalism style     |

## Feedback

`POST /feedback` with `{ suggestionId, action: "accept" | "reject", comment? }`
(`comment` up to 4,000 characters) returns 202. The hub already knows the
suggestion by its id. The add-in never waits for or shows the outcome: every
failure is logged and swallowed so it cannot block resolving suggestions.

## Correction Preferences

`GET /preferences` and `PUT /preferences { correctionInstructions: string | null }`
both return `{ correctionInstructions, correctionInstructionsMaxLength }`.
The analysis profile stays add-in state in `OfficeRuntime.storage`.
