# API Contract — Backend Requirements

This document specifies what the Stylistic frontend (Word add-in) expects from the Mastra backend. The backend team must implement a Mastra workflow that conforms to this contract.

## Technology

- **Framework:** [Mastra](https://mastra.ai/) (TypeScript AI framework)
- **Communication:** The frontend uses `@mastra/client-js` v1.7.1 to call the workflow
- **Protocol:** Mastra's built-in HTTP API for workflows, plus Better Auth routes for login/session management
- **Authentication:** Better Auth bearer session token sent as `Authorization: Bearer <token>` on Mastra workflow calls

## Workflow Definitions

### 1. `stylistic-workflow`

```
stylistic-workflow
```

The frontend calls this workflow via:

```typescript
const workflow = client.getWorkflow("stylistic-workflow");
const run = await workflow.createRun();
await run.start({ inputData: { ... } });
const { runId } = run;
const result = await workflow.runById(runId, {
  fields: ["result", "error"],
  withNestedWorkflows: false,
});
```

In `@mastra/client-js` v1.7.1, async submission is implemented as `createRun()` + `run.start()` + later `workflow.runById(runId)`. The frontend submits each chunk, stores `run.runId`, and polls `workflow.runById(runId)` until the run reaches a supported terminal state.

Important: in this SDK version, `status` is always included in the response metadata and must NOT be requested inside `fields`. The frontend only requests explicit payload fields such as `result` and `error`; adding `status` to `fields` causes a deterministic HTTP 400.

### Mastra Server

- **Default URL:** `http://localhost:4111` (Mastra's default port)
- **CORS:** Must allow requests from `https://localhost:3000` (add-in dev server)

### Authentication Contract

Before calling the workflow, the taskpane must have a valid Better Auth session.
The backend is expected to expose:

| Route                  | Purpose                                                     |
| ---------------------- | ----------------------------------------------------------- |
| `/auth/*`              | Better Auth sign-in, callback, session, and logout routes.  |
| `/auth-complete`       | Backend OAuth completion bridge used after Google callback. |
| `/auth-bridge-session` | One-time code exchange used by the Office Dialog page.      |

The add-in stores the Better Auth session token with `OfficeRuntime.storage` and
creates Mastra clients with the current bearer token. If the backend returns 401
for a protected workflow call, the taskpane should clear the local session and
require the user to sign in again.

The add-in also stores the user-selected analysis profile in
`OfficeRuntime.storage`, but that preference remains frontend-owned state. The
selection is validated against the frontend's supported `AnalysisProfileId`
whitelist during bootstrap, then forwarded as `genero` in each workflow call.
The backend should trust the incoming `genero` value; it does not need a
separate preference endpoint for profile persistence.

The add-in must not store Google provider access/refresh tokens. Provider tokens
belong to the backend/auth provider boundary; the frontend only uses the Better
Auth session token.

## Input Schema

The analysis workflow must accept this `inputData`:

```typescript
interface WorkflowInput {
  /** Text to analyze (up to ~100K characters). */
  text: string;

  /** Stable document UUID generated and persisted by the add-in. */
  documentUuid: string;

  /** Editorial genre for analysis style. */
  genero?: "narrativa-literaria" | "ensayo-academico" | "periodismo-cultural" | "general";

  /** Optional document title shown/stored by the backend. */
  title?: string;

  /** Optional document-level processing preferences. */
  processingConfig?: Record<string, unknown>;
}
```

### Example Input

```json
{
  "text": "Básicamente, es completamente necesario utilizar este periodo de tiempo con el objetivo de realizar la tarea.",
  "documentUuid": "11111111-1111-4111-8111-111111111111",
  "genero": "general",
  "title": "Draft chapter 03",
  "processingConfig": {
    "chunking": {
      "mode": "default"
    }
  },
}
```

### Notes on Input Fields

- **`text`** — The frontend chunks large documents at paragraph boundaries and sends each chunk as a separate workflow execution. The backend receives plain text and should not assume anything about document structure.
- **`documentUuid`** — Stable add-in-generated UUID that identifies the Word document across analysis and feedback flows. The add-in persists it in `Office.context.document.settings` and reuses it across runs.
- **`genero`** — Optional document-level editorial style. Supported values: `narrativa-literaria`, `ensayo-academico`, `periodismo-cultural`, `general`. Backend default remains `general`.
- **`title`** — Optional document title metadata sent by the add-in.
- **`processingConfig`** — Optional document-level processing preferences that the backend stores with the document context.

### Contract Rules

- Normal analysis must enter through `stylistic-workflow`.
- The add-in must **not** call `POST /documents/resolve` before analysis. The backend resolves document context and auto-upserts persisted document data inside the workflow when `documentUuid` is present.
- The frontend sends `documentUuid` on every analysis request; `autorSlug` is no longer part of the add-in contract.
- The contract is document-based, not author-slug-based.

## Output Format

When polling returns `status: "success"`, `result.result` must conform to:

```typescript
interface WorkflowOutput {
  /** Array of editorial suggestions for the analyzed text. */
  suggestions: Array<{
    /** Paragraph-level context used to locate the suggestion in the document. */
    context: string;

    /** Exact substring within `context` targeted by the suggestion. */
    anchor: string;

    /**
     * Transport text for "track-change" suggestions.
     *
     * Required for "track-change" and optional for "comment-only".
     * For track-change:
     * - plain non-empty text means replace anchor with that text,
     * - empty string means delete-only,
     * - exact markdown `*anchor*` / `**anchor**` means italic/bold formatting.
     */
    suggestedText?: string;

    /** Human-readable justification shown to the user. */
    justification: string;

    /** Category label (e.g., "Redundancia", "Muletilla", "Elección de palabra"). */
    category: string;

    /** How critical the suggestion is. */
    severity: "high" | "medium" | "low";

    /** Suggestion type. Defaults to "track-change" if not specified. */
    type?: "track-change" | "comment-only";
  }>;

  /** Optional warnings from the backend (e.g., "text too short for meaningful analysis"). */
  warnings?: string[];
}
```

### Example Output

```json
{
  "suggestions": [
    {
      "context": "Básicamente, es completamente necesario utilizar este periodo de tiempo.",
      "anchor": "Básicamente, ",
      "suggestedText": "",
      "justification": "Muletilla que debilita la afirmación.",
      "category": "Muletilla",
      "severity": "medium",
      "type": "track-change"
    },
    {
      "context": "Era completamente necesario terminarlo hoy.",
      "anchor": "completamente necesario",
      "suggestedText": "necesario",
      "justification": "\"Necesario\" ya implica completitud.",
      "category": "Redundancia",
      "severity": "high",
      "type": "track-change"
    },
    {
      "context": "El periodo de tiempo era necesario.",
      "anchor": "periodo de tiempo",
      "suggestedText": "periodo",
      "justification": "\"Periodo\" ya denota tiempo.",
      "category": "Redundancia",
      "severity": "high",
      "type": "track-change"
    },
    {
      "context": "Ese era el inicio del post mortem reportado por PRIME.",
      "anchor": "post mortem",
      "suggestedText": "*post mortem*",
      "justification": "Locución latina que debe ir en cursiva.",
      "category": "Formato tipográfico",
      "severity": "medium",
      "type": "track-change"
    },
    {
      "context": "El informe fue marcado por PRIME.",
      "anchor": "PRIME",
      "suggestedText": "**PRIME**",
      "justification": "Sigla editorial que debe destacarse en negrita.",
      "category": "Formato tipográfico",
      "severity": "low",
      "type": "track-change"
    },
    {
      "context": "Utilizar este periodo de tiempo con el objetivo de realizar la tarea.",
      "anchor": "con el objetivo de",
      "suggestedText": "para",
      "justification": "\"Para\" es más directo.",
      "category": "Elección de palabra",
      "severity": "medium",
      "type": "track-change"
    }
  ]
}
```

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

## Genre (genero) Behavior

The `genero` field tells the workflow what analysis style to apply:

| Genre                 | Description                   |
| --------------------- | ----------------------------- |
| `general`             | Everyday writing improvements |
| `narrativa-literaria` | Literary narrative style      |
| `ensayo-academico`    | Academic essay style          |
| `periodismo-cultural` | Cultural journalism style     |

## Error Handling

The frontend handles workflow failures gracefully during explicit submit/poll:

- **Submit without `runId`** — The frontend retries chunk submission up to 3 times.
- **Submit with `runId`** — The frontend assumes the run was accepted and switches to polling.
- **Polling `status: "running" | "pending" | "waiting"`** — Not an error; the frontend keeps polling.
- **Polling `status: "suspended" | "paused"`** — Treated as a failed chunk because the run would require `resume()` and this frontend does not implement resume flows.
- **Polling `status: "failed" | "tripwire" | "canceled" | "bailed"`** — Treated as a failed chunk.
- **Polling network errors / thrown errors** — The frontend retries the poll up to 3 times, then marks the chunk as failed.
- **Empty suggestions array** — Treated as "no issues found" (valid response).
- **Malformed output** — Treated as a failed chunk (suggestions array missing or not an array).

The backend does **not** need to implement error responses — Mastra handles HTTP error codes automatically.

## Text Size Expectations

The frontend splits documents into chunks of up to 100,000 characters. The backend should expect:

| Chunk Size     | Words (~)    | Processing Time |
| -------------- | ------------ | --------------- |
| < 5K chars     | < 800 words  | < 10s           |
| 5K–30K chars   | 800–5K words | 10–30s          |
| 30K–100K chars | 5K–16K words | 30–120s         |

The backend should be able to handle 100K characters in a single workflow execution without timing out. If the backend's AI model has a smaller context window, the workflow should handle internal sub-chunking transparently.

---

## 2. `feedback-workflow`

Explicit suggestion feedback is fire-and-forget from the add-in point of view and uses the same document-scoped identity model as analysis.

The frontend calls this workflow via:

```typescript
const workflow = client.getWorkflow("feedback-workflow");
const run = await workflow.createRun();
await run.start({ inputData: payload });
```

The add-in does not poll for a semantic result today. Feedback must remain non-blocking for UX.

### Input Schema

```typescript
interface FeedbackWorkflowInput {
  /** Stable document UUID generated and persisted by the add-in. */
  documentUuid: string;

  category: string;
  context: string;
  anchor: string;
  suggestedText?: string;
  justification: string;
  action: "accept" | "reject";
  severity: "high" | "medium" | "low";
  suggestionType: "track-change" | "comment-only";
  comment?: string;
}
```

### Feedback Contract Rules

- The payload key is `documentUuid`.
- The backend updates persisted `document_style_profile` data using that `documentUuid`.
- Feedback remains fire-and-forget and must never block the resolution UX.
- `FeedbackPayload`, `SuggestionResolutionWorkflow`, and `FeedbackAdapter` must stay aligned with this document-based contract.
