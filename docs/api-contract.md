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
const result = await run.start({ inputData: { ... } });
```

### Mastra Server

- **Default URL:** `http://localhost:4111` (Mastra's default port)
- **CORS:** Must allow requests from `https://localhost:3000` (add-in dev server)

## Input Schema

The workflow must accept this `inputData`:

```typescript
interface WorkflowInput {
  /** Text chunk to analyze (up to ~100K characters). */
  text: string;

  /** Analysis profile: "general", "formal", or "academic". */
  profile: string;

  /** Zero-based index of this chunk within the full document. */
  chunkIndex: number;

  /** Total number of chunks the document was split into. */
  totalChunks: number;
}
```

### Example Input

```json
{
  "text": "Básicamente, es completamente necesario utilizar este periodo de tiempo con el objetivo de realizar la tarea.",
  "profile": "formal",
  "chunkIndex": 0,
  "totalChunks": 1
}
```

## Output Format

On `status: "success"`, `result.result` must conform to:

```typescript
interface WorkflowOutput {
  suggestions: Array<{
    /** EXACT substring from the input text (case-sensitive). */
    originalText: string;

    /** Replacement text. */
    suggestedText: string;

    /** Human-readable justification shown to the user. */
    justification: string;

    /** Category label (e.g., "Redundancia", "Muletilla", "Elección de palabra"). */
    category: string;
  }>;
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
      "category": "Muletilla"
    },
    {
      "originalText": "completamente necesario",
      "suggestedText": "necesario",
      "justification": "\"Necesario\" ya implica completitud.",
      "category": "Redundancia"
    },
    {
      "originalText": "utilizar",
      "suggestedText": "usar",
      "justification": "\"Usar\" es más simple y directo.",
      "category": "Elección de palabra"
    },
    {
      "originalText": "periodo de tiempo",
      "suggestedText": "periodo",
      "justification": "\"Periodo\" ya denota tiempo.",
      "category": "Redundancia"
    },
    {
      "originalText": "con el objetivo de",
      "suggestedText": "para",
      "justification": "\"Para\" es más directo.",
      "category": "Elección de palabra"
    }
  ]
}
```

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

The backend may adjust the AI model's instructions based on the profile. If the backend doesn't differentiate between profiles, it can ignore the field — the frontend handles this gracefully.

## Error Handling

The frontend handles workflow failures gracefully:

- **`status: "failed"`** — The frontend retries the chunk up to 3 times, then skips it.
- **Network errors** — Same retry behavior.
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

The current version analyzes **Spanish text only**. The `text` field will always contain Spanish content. Justification messages should also be in Spanish.
