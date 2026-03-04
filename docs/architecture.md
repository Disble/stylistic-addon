# Architecture

This document describes the layered architecture of Stylistic, the data flow between modules, and the key design decisions behind them.

## Design Principles

1. **Separation of concerns** — each module owns a single responsibility. The UI never calls Word directly. The Word API layer never calls the backend. The backend client never touches the DOM.
2. **No frontend business logic** — all text analysis is performed server-side by a Mastra workflow. The frontend is purely UI/UX + transport.
3. **Preserve-and-restore** — the add-in always reads the document's `changeTrackingMode` before modifying it and restores it when done, even if an error occurs (via `try/finally`).
4. **Batched operations** — suggestions are applied in groups via separate `Word.run` calls, each independently committed. This prevents a single failure from losing all progress.
5. **Reliability over speed** — every chunk and batch has retry logic. Partial failures are reported, not fatal. The user gets as many suggestions as possible.

## System Overview

```
┌─────────────────────────────────────────────────────┐
│                   Word Add-in (Frontend)            │
│                                                     │
│  ┌───────────────────────────────────────────────┐  │
│  │  taskpane.ts — UI orchestrator                │  │
│  │  Events, progress, rendering                  │  │
│  └──────┬──────────────┬──────────────┬──────────┘  │
│         │              │              │              │
│  ┌──────▼──────┐ ┌─────▼──────┐ ┌────▼─────┐       │
│  │ wordApi.ts  │ │mastraClient│ │chunker.ts│       │
│  │ Office.js   │ │ .ts        │ │ text     │       │
│  │ read/write  │ │ workflow   │ │ splitting│       │
│  └──────┬──────┘ │ execution  │ └──────────┘       │
│         │        └─────┬──────┘                     │
│         ▼              │                             │
│  ┌────────────┐        │                             │
│  │ Word       │        │  HTTP (Mastra client-js)    │
│  │ Document   │        │                             │
│  └────────────┘        ▼                             │
│                 ┌──────────────┐                     │
│                 │ Mastra Server│                     │
│                 │ (Backend)    │                     │
│                 └──────────────┘                     │
└─────────────────────────────────────────────────────┘
```

## Module Map

```
src/
├── lib/
│   ├── types.ts          ← Shared interfaces (no runtime code)
│   ├── config.ts         ← Constants (URLs, batch sizes, retry policy)
│   ├── wordApi.ts        ← Word API abstraction (only Office.js consumer)
│   ├── mastraClient.ts   ← Mastra workflow client (@mastra/client-js wrapper)
│   └── chunker.ts        ← Text chunking (pure function, paragraph-boundary)
└── taskpane/
    ├── taskpane.ts       ← UI orchestrator (event binding, progress, rendering)
    ├── taskpane.html     ← Task pane markup
    └── taskpane.css      ← Styling (Fluent UI compatible)
```

### Dependency Graph

```
taskpane.ts
├── imports wordApi.ts
├── imports mastraClient.ts
├── imports chunker.ts
├── imports config.ts
└── imports types.ts

wordApi.ts
├── imports types.ts
├── imports config.ts
└── uses global: Word (Office.js)

mastraClient.ts
├── imports types.ts
├── imports config.ts
└── imports @mastra/client-js

chunker.ts
├── imports types.ts
└── imports config.ts

config.ts
└── imports types.ts

types.ts
└── (no imports)
```

The dependency graph is strictly acyclic. `types.ts` is at the bottom. `wordApi.ts`, `mastraClient.ts`, and `chunker.ts` are peers that never import each other. `taskpane.ts` is the composition root.

## Data Flow

The main analysis flow follows this sequence:

```
User clicks "Analizar y sugerir"
        │
        ▼
taskpane.ts: handleAnalyze()
        │
        ├──► wordApi.getDocumentText()
        │         └── Word.run → body.load("text") → context.sync()
        │         └── returns plain text string
        │
        ├──► mastraClient.checkConnection()
        │         └── client.getWorkflow() → workflow.details()
        │         └── returns boolean (fail-fast gate)
        │
        ├──► chunker.splitText(text, maxChunkSize)
        │         └── splits at paragraph boundaries
        │         └── returns TextChunk[]
        │
        ├──► For each chunk (sequential):
        │     └── mastraClient.analyzeChunk(chunk, profile)
        │           └── workflow.createRun() → run.start({ inputData })
        │           └── retry on failure (up to 3 times, exponential backoff)
        │           └── returns ChunkResult { suggestions[], error? }
        │
        ├──► deduplicateByOriginalText(allSuggestions)
        │         └── removes cross-chunk duplicates (case-insensitive)
        │
        ├──► wordApi.applySuggestionsInBatches(suggestions, onProgress)
        │         │
        │         ├── Word.run #1: save + set TrackAll
        │         ├── Word.run #2..N: for each batch of 30:
        │         │     ├── Enqueue body.search() for each suggestion
        │         │     ├── context.sync() (one round-trip)
        │         │     ├── insertText(replace) for found matches
        │         │     ├── context.sync() (commit)
        │         │     └── report progress via callback
        │         └── Word.run #final (finally): restore tracking mode
        │         └── returns InsertionResult
        │
        └──► taskpane.ts: renderResults(suggestions, result, chunkErrors)
                  └── Displays applied/failed suggestions in the task pane
```

## Key Design Decisions

### Why a Mastra workflow instead of a local analyzer?

The PoC used a local regex-based analyzer (`analyzer.ts`). The production version uses an AI-powered backend for:
- **Semantic understanding** — AI can detect context-dependent issues that regex cannot.
- **Extensibility** — new analysis rules are prompt changes, not code changes.
- **Separation of concerns** — the frontend doesn't need to know about NLP, models, or prompts.

### Why `@mastra/client-js` instead of raw `fetch`?

Mastra provides a typed SDK that handles workflow execution, run management, and the Mastra HTTP protocol. Using raw `fetch` would require reimplementing the workflow run lifecycle (create run → start → poll for result).

### Why chunk on the frontend?

The backend's AI model has a context window limit. Rather than sending a 1.2 MB document and hoping the backend handles it, the frontend:
1. Chunks at paragraph boundaries (preserving semantic context).
2. Sends chunks sequentially (one workflow execution per chunk).
3. Retries individual chunks on failure (not the entire document).
4. Reports chunk-level progress to the user.

### Why separate Word.run per batch?

Each `Word.run` is an independent transaction. Changes from `Word.run #3` are committed to the document before `Word.run #4` starts. If `#4` fails:
- Suggestions from batches 1–3 are already saved as tracked changes.
- The user doesn't lose progress.
- The failure is reported, not catastrophic.

The alternative (one `Word.run` for all suggestions) risks losing everything if it fails mid-execution.

### Why `start()` instead of `stream()` for workflows?

`run.start()` waits for the complete result. `run.stream()` provides real-time events. We chose `start()` because:
- We need the complete suggestion array before applying Track Changes.
- Progress comes from chunk-level iteration, not intra-workflow events.
- Retry logic is simpler with a single request-response per chunk.

### Why deduplicate in the orchestrator?

Multiple chunks may contain the same phrase. Without deduplication, the second `body.search()` for the same `originalText` would find the already-replaced text and fail. Deduplication is a data integrity measure (preventing duplicate Track Changes), not business logic.

## Error Handling Strategy

Errors are handled at four levels:

| Level | Module | Strategy |
|---|---|---|
| Word API | `wordApi.ts` | `try/finally` ensures tracking mode is always restored. Each batch is independent. |
| Backend | `mastraClient.ts` | Retry with exponential backoff (3 attempts). Never throws — returns empty result on failure. |
| Orchestrator | `taskpane.ts` | `try/catch` around the full pipeline. Errors translated via `toUserMessage()`. Partial results are preserved. |
| UI | `taskpane.ts` | `setAnalyzeLoading(false)` runs in `finally`. Progress bar resets on completion. |

### Partial Success Philosophy

The system is designed for partial success:
- If 3/5 chunks succeed → suggestions from those 3 chunks are applied.
- If 25/30 suggestions in a batch are found → those 25 are applied, 5 are reported as "not found".
- The results panel always shows the complete picture: applied count, failed count, and chunk errors.
