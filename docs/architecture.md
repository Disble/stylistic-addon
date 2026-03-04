# Architecture

This document describes the layered architecture of Stylistic, the data flow between modules, and the key design decisions behind them.

## Design Principles

1. **Separation of concerns** — each module owns a single responsibility. The UI never calls Word directly. The Word API layer never calls the backend. The backend client never touches the DOM.
2. **No frontend business logic** — all text analysis is performed server-side by a Mastra workflow. The frontend is purely UI/UX + transport.
3. **Preserve-and-restore** — the add-in always reads the document's `changeTrackingMode` before modifying it and restores it when done, even if an error occurs (via `try/finally`).
4. **Per-suggestion isolation** — each suggestion is applied in its own `Word.run` call. This prevents stale ranges after OOXML insertions shift text, and ensures partial failures don't lose already-applied changes.
5. **Reliability over speed** — every chunk has retry logic. Partial failures are reported, not fatal. The user gets as many suggestions as possible.

## System Overview

```
┌─────────────────────────────────────────────────┐
│  Word Add-in (Frontend)                         │
│                                                 │
│  taskpane.ts — UI orchestrator                  │
│  ├── wordApi.ts    — Read document, apply       │
│  │                   Track Changes via OOXML    │
│  ├── mastraClient.ts — Workflow execution       │
│  │                     via @mastra/client-js     │
│  └── chunker.ts    — Split large texts at       │
│                      paragraph boundaries       │
├─────────────────────────────────────────────────┤
│  Mastra Backend (separate application)          │
│  └── editorial-workflow — AI text analysis      │
└─────────────────────────────────────────────────┘
```

| Layer | Module | Responsibility | Imports Office.js? |
|---|---|---|---|
| UI | `taskpane.ts` | Event handling, progress, rendering | No (delegates) |
| API | `wordApi.ts` | Document read/write, OOXML tracked changes, comment cleanup | Yes (only module) |
| Backend | `mastraClient.ts` | Workflow execution, retry logic | No |
| Chunker | `chunker.ts` | Text splitting at paragraph boundaries | No |
| Config | `config.ts` | Internal constants and defaults | No |
| Types | `types.ts` | Shared interfaces | No |

## Module Map

```
src/
├── lib/
│   ├── types.ts          ← Shared interfaces (no runtime code)
│   ├── config.ts         ← Constants (URLs, retry policy)
│   ├── wordApi.ts        ← Word API abstraction (OOXML tracked changes + cleanup)
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
        │     └── mastraClient.analyzeChunk(chunk, profile, language)
        │           └── workflow.createRun() → run.start({ inputData })
        │           └── retry on failure (up to 3 times, exponential backoff)
        │           └── returns ChunkResult { suggestions[], error? }
        │
        ├──► deduplicateByOriginalText(allSuggestions)
        │         └── removes cross-chunk duplicates (case-insensitive)
        │
        ├──► wordApi.applySuggestionsInBatches(suggestions, onProgress)
        │         │
        │         └── For each suggestion (one Word.run per suggestion):
        │               ├── body.search(originalText, matchCase)
        │               ├── Extract <w:rPr> formatting from matched range
        │               ├── Disable changeTrackingMode
        │               ├── Build OOXML package:
        │               │     ├── <w:del> with original text
        │               │     ├── <w:ins> with replacement text
        │               │     └── <w:comment> with [Category] + justification
        │               ├── range.insertOoxml(package, replace)
        │               ├── Restore changeTrackingMode (in finally block)
        │               └── Report progress via callback
        │         └── returns InsertionResult
        │
        ├──► Shows "Limpiar comentarios resueltos" button
        │
        └──► taskpane.ts: renderResults(suggestions, result, chunkErrors)
                  └── Displays applied/failed suggestions in the task pane
```

### Comment Cleanup Flow

After the user accepts or rejects tracked changes in Word, comments remain orphaned. The cleanup flow:

```
User clicks "Limpiar comentarios resueltos"
        │
        ▼
taskpane.ts: handleCleanup()
        │
        └──► wordApi.cleanupResolvedComments()
                  │
                  ├── Sync 1: Load all tracked changes (author, type)
                  │           and comments (authorName) from the document
                  │
                  ├── Filter to Stylistic-authored items only
                  │
                  ├── Sync 2: Get document ranges for each comment
                  │           and tracked change via getRange()
                  │
                  ├── Sync 3: Compare every comment range against
                  │           every TC range via compareLocationWith()
                  │
                  ├── For each Stylistic comment:
                  │     └── If no TC overlaps → orphaned → delete()
                  │     └── If TC overlaps → still pending → keep
                  │
                  └── Sync 4: Execute deletes
                  └── returns { deleted, kept }
```

## OOXML Strategy

All changes are applied via flat OPC OOXML packages rather than Word's `insertText()` API. Each package contains 4 parts:

1. **`/_rels/.rels`** — Package relationships (points to document.xml)
2. **`/word/_rels/document.xml.rels`** — Document relationships (points to comments.xml)
3. **`/word/document.xml`** — Tracked change markup (`<w:del>` + `<w:ins>`) with comment anchors
4. **`/word/comments.xml`** — Formatted justification (bold category + justification text)

This approach ensures:
- The tracked change blue card shows `w:author="Stylistic"` cleanly
- The comment appears as a margin balloon with formatting
- Original text formatting (`<w:rPr>`) is preserved in the tracked change
- No double-tracking: the document's `changeTrackingMode` is temporarily disabled during OOXML insertion

## Key Design Decisions

### Why a Mastra workflow instead of a local analyzer?

The production version uses an AI-powered backend for:
- **Semantic understanding** — AI can detect context-dependent issues that regex cannot.
- **Extensibility** — new analysis rules are prompt changes, not code changes.
- **Separation of concerns** — the frontend doesn't need to know about NLP, models, or prompts.

### Why `@mastra/client-js` instead of raw `fetch`?

Mastra provides a typed SDK that handles workflow execution, run management, and the Mastra HTTP protocol. Using raw `fetch` would require reimplementing the workflow run lifecycle (create run → start → poll for result).

### Why chunk on the frontend?

The backend's AI model has a context window limit. Rather than sending a large document and hoping the backend handles it, the frontend:
1. Chunks at paragraph boundaries (preserving semantic context).
2. Sends chunks sequentially (one workflow execution per chunk).
3. Retries individual chunks on failure (not the entire document).
4. Reports chunk-level progress to the user.

### Why one Word.run per suggestion?

Each suggestion is applied in its own `Word.run` to avoid stale ranges. After an OOXML insertion replaces a range, all subsequent ranges in the document may shift. By using a fresh `Word.run` per suggestion:
- Each search starts from a clean document state
- A failure in suggestion N doesn't affect suggestions 1..(N-1)
- Progress is reported after each individual suggestion

### Why `start()` instead of `stream()` for workflows?

`run.start()` waits for the complete result. `run.stream()` provides real-time events. We chose `start()` because:
- We need the complete suggestion array before applying Track Changes.
- Progress comes from chunk-level iteration, not intra-workflow events.
- Retry logic is simpler with a single request-response per chunk.

### Why deduplicate in the orchestrator?

Multiple chunks may contain the same phrase. Without deduplication, the second `body.search()` for the same `originalText` would find the already-replaced text and fail. Deduplication is a data integrity measure (preventing duplicate Track Changes), not business logic.

### Why range colocation for comment cleanup?

Comments and tracked changes have no direct link in the Word API. The cleanup uses `Comment.getRange()` and `TrackedChange.getRange()` with `Range.compareLocationWith()` to determine if a comment is still anchored to a pending tracked change. This approach:
- Uses the document as the source of truth (no in-memory state)
- Works across sessions (no registry to persist)
- Never deletes comments it can't positively identify as orphaned

## Error Handling Strategy

Errors are handled at four levels:

| Level | Module | Strategy |
|---|---|---|
| Word API | `wordApi.ts` | `try/finally` ensures tracking mode is always restored. Each suggestion is independent. |
| Backend | `mastraClient.ts` | Retry with exponential backoff (3 attempts). Never throws — returns empty result on failure. |
| Orchestrator | `taskpane.ts` | `try/catch` around the full pipeline. Errors translated via `toUserMessage()`. Partial results are preserved. |
| UI | `taskpane.ts` | `setAnalyzeLoading(false)` runs in `finally`. Progress bar resets on completion. |

### Partial Success Philosophy

The system is designed for partial success:
- If 3/5 chunks succeed → suggestions from those 3 chunks are applied.
- If 25/30 suggestions are found → those 25 are applied, 5 are reported as "not found".
- The results panel always shows the complete picture: applied count, failed count, and chunk errors.
