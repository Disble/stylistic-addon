# API Contract — Backend Requirements

This document specifies what the Stylistic frontend (Word add-in) expects from the Mastra backend. The backend team must implement a Mastra workflow that conforms to this contract.

## Technology

- **Framework:** [Mastra](https://mastra.ai/) (TypeScript AI framework)
- **Communication:** The frontend uses `@mastra/client-js` v1.7.1 to call the workflow
- **Protocol:** Mastra's built-in HTTP API (no custom REST endpoints needed)

## Workflow Definition

### Workflow ID

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

## Input Schema

The workflow must accept this `inputData`:

```typescript
interface WorkflowInput {
  /** Text to analyze (up to ~100K characters). */
  text: string;

  /** Editorial genre for analysis style. */
  genero: "narrativa-literaria" | "ensayo-academico" | "periodismo-cultural" | "general";

  /** Author slug for personalization and author tracking. */
  autorSlug: string;
}
```

### Example Input

```json
{
  "text": "Básicamente, es completamente necesario utilizar este periodo de tiempo con el objetivo de realizar la tarea.",
  "genero": "general",
  "autorSlug": "Disble"
}
```

### Notes on Input Fields

- **`text`** — The frontend chunks large documents at paragraph boundaries and sends each chunk as a separate workflow execution. The backend receives plain text and should not assume anything about document structure.
- **`genero`** — Determines the editorial analysis style. Supported values: `narrativa-literaria`, `ensayo-academico`, `periodismo-cultural`, `general`.
- **`autorSlug`** — Author identifier for personalization and author tracking. Currently defaults to `"Disble"`. Backend can use this to maintain author-specific writing profiles.

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

    /** Replacement text. Required for "track-change", optional for "comment-only". */
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

| Severity | When to use | Example |
|---|---|---|
| `high` | Clear errors, redundancies, grammatical issues | "periodo de tiempo" → "periodo" |
| `medium` | Stylistic improvements, filler words | "Básicamente, " → "" |
| `low` | Minor preferences, optional simplifications | "utilizar" → "usar" |

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

**Don't:**
- Return `"Completamente necesario"` (wrong capitalization)
- Return `"completamente  necesario"` (extra space)
- Return an anchor that is not contained in context

## Genre (genero) Behavior

The `genero` field tells the workflow what analysis style to apply:

| Genre | Description |
|---|---|
| `general` | Everyday writing improvements |
| `narrativa-literaria` | Literary narrative style |
| `ensayo-academico` | Academic essay style |
| `periodismo-cultural` | Cultural journalism style |

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

| Chunk Size | Words (~) | Processing Time |
|---|---|---|
| < 5K chars | < 800 words | < 10s |
| 5K–30K chars | 800–5K words | 10–30s |
| 30K–100K chars | 5K–16K words | 30–120s |

The backend should be able to handle 100K characters in a single workflow execution without timing out. If the backend's AI model has a smaller context window, the workflow should handle internal sub-chunking transparently.
