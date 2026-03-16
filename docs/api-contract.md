# API Contract — Backend Requirements

This document specifies what the Stylistic frontend (Word add-in) expects from the Mastra backend. The backend team must implement a Mastra workflow that conforms to this contract.

## Technology

- **Framework:** [Mastra](https://mastra.ai/) (TypeScript AI framework)
- **Communication:** The frontend uses `@mastra/client-js` v1.7.1 to call the workflow
- **Protocol:** Mastra's built-in HTTP API (no custom REST endpoints needed)

## Workflow Definition

### Workflow ID

```
editorial-workflow
```

The frontend calls this workflow via:

```typescript
const workflow = client.getWorkflow("editorial-workflow");
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

  /** Analysis profile: "general", "formal", or "academic". */
  profile: string;

  /** ISO 639-1 language code of the text (e.g., "es", "en"). */
  language: string;
}
```

### Example Input

```json
{
  "text": "Básicamente, es completamente necesario utilizar este periodo de tiempo con el objetivo de realizar la tarea.",
  "profile": "formal",
  "language": "es"
}
```

### Notes on Input Fields

- **`text`** — The frontend chunks large documents at paragraph boundaries and sends each chunk as a separate workflow execution. The backend receives plain text and should not assume anything about document structure.
- **`profile`** — Determines the editorial analysis style. If the backend doesn't differentiate between profiles, it can ignore the field — the frontend handles this gracefully.
- **`language`** — The text's language. Currently the frontend always sends `"es"` (Spanish). The backend should use this to select the appropriate editorial rules and justification language.

## Output Format

When polling returns `status: "success"`, `result.result` must conform to:

```typescript
interface WorkflowOutput {
  /** Array of editorial suggestions for the analyzed text. */
  suggestions: Array<{
    /** EXACT substring from the input text (case-sensitive). */
    originalText: string;

    /** Replacement text. */
    suggestedText: string;

    /** Human-readable justification shown to the user. */
    justification: string;

    /** Category label (e.g., "Redundancia", "Muletilla", "Elección de palabra"). */
    category: string;

    /** How critical the suggestion is. */
    severity: "high" | "medium" | "low";
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
      "originalText": "Básicamente, ",
      "suggestedText": "",
      "justification": "Muletilla que debilita la afirmación.",
      "category": "Muletilla",
      "severity": "medium"
    },
    {
      "originalText": "completamente necesario",
      "suggestedText": "necesario",
      "justification": "\"Necesario\" ya implica completitud.",
      "category": "Redundancia",
      "severity": "high"
    },
    {
      "originalText": "utilizar",
      "suggestedText": "usar",
      "justification": "\"Usar\" es más simple y directo.",
      "category": "Elección de palabra",
      "severity": "low"
    },
    {
      "originalText": "periodo de tiempo",
      "suggestedText": "periodo",
      "justification": "\"Periodo\" ya denota tiempo.",
      "category": "Redundancia",
      "severity": "high"
    },
    {
      "originalText": "con el objetivo de",
      "suggestedText": "para",
      "justification": "\"Para\" es más directo.",
      "category": "Elección de palabra",
      "severity": "medium"
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

## Critical Constraint: Exact Substring Matching

The `originalText` in every suggestion **must** be an exact, character-for-character substring of the input `text`. The frontend uses Word's `body.search()` API with `matchCase: true` to locate the text. If the substring doesn't match exactly:

- The suggestion will silently fail (reported as "not found")
- No tracked change will be inserted for that suggestion

**Do:**
- Return `"completamente necesario"` (exact match from input)
- Include trailing punctuation/spaces if they're part of the replacement (e.g., `"Básicamente, "` to remove the comma and space)

**Don't:**
- Return `"Completamente necesario"` (wrong capitalization)
- Return `"completamente  necesario"` (extra space)
- Return approximate or paraphrased text

## Profile Behavior

The `profile` field tells the workflow what analysis style to apply:

| Profile | Description | Aggressiveness |
|---|---|---|
| `general` | Everyday writing improvements | Moderate |
| `formal` | Business and professional tone | Conservative |
| `academic` | Academic and technical writing | Conservative, respects domain terms |

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

## Language

The `language` field indicates the text's language as an ISO 639-1 code. Currently the frontend sends `"es"` (Spanish). Justification messages should match the text language (Spanish justifications for Spanish text).
