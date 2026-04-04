# Architecture

This document describes the architecture of Stylistic, the design patterns used, the data flow, and the rationale behind every key decision.

> **Important correction:** this document explains the current architecture and its patterns, but the authoritative clarification of the addon's frontend domain and the new Track Changes lifecycle policy now lives in [`review-domain-and-track-changes.md`](./review-domain-and-track-changes.md).

---

## Frontend Domain Clarification

The `stylistic-addon` frontend domain is:

> **Presentar al usuario sugerencias de estilo provenientes del backend y permitirle aceptarlas, rechazarlas o debatirlas dentro de Word.**

This clarification matters because several implementation mechanisms are important but are **not** the full frontend domain on their own:

- `DocumentReviewState` is an important internal support concept, but not the whole domain.
- `PipelineStateMachine` and the analysis pipeline are workflow mechanisms, not the domain definition.
- `changeTrackingMode` and Word content controls are infrastructure/host concerns, not domain language.

The current documentation should therefore be read with this distinction in mind:

- **domain goal** → user interaction with style suggestions in Word,
- **supporting state** → `DocumentReviewState`, pending/resolved review artifacts,
- **workflow state** → `PipelineStateMachine`, future resolution workflow state,
- **infrastructure** → Office.js, tags, comments, polling, `changeTrackingMode`.

---

## Architectural Pattern: Hexagonal Architecture (Ports & Adapters)

Stylistic follows a **Hexagonal Architecture** (also known as Ports & Adapters), adapted for a TypeScript browser add-in. The dependency rule is strict: outer layers depend on inner layers; inner layers never depend on outer layers.

```
┌─────────────────────────────────────────────────────────────────┐
│  PRESENTATION (taskpane/)                                       │
│  taskpane.ts — event binding, observer registration, render     │
├─────────────────────────────────────────────────────────────────┤
│  APPLICATION / DOMAIN (domain/)                                 │
│  pipeline/ — Chain of Responsibility orchestrator + 7 handlers  │
│  ports.ts  — IDocumentPort, IAnalysisPort (the port contracts)  │
│  types.ts  — Shared interfaces (zero runtime code)             │
├────────────────────────┬────────────────────────────────────────┤
│  ADAPTER: Word         │  ADAPTER: Mastra                       │
│  adapters/word/        │  adapters/mastra/ + RetryDecorator     │
│  (implements           │  (implements                           │
│   IDocumentPort)       │   IAnalysisPort)                       │
├────────────────────────┴────────────────────────────────────────┤
│  INFRASTRUCTURE (infrastructure/)                               │
│  config.ts — constants         chunker.ts — pure text splitting │
└─────────────────────────────────────────────────────────────────┘
```

### Why Hexagonal?

| Goal | How the architecture achieves it |
|------|----------------------------------|
| **Testability** | The pipeline depends only on `IDocumentPort` and `IAnalysisPort`. Mock implementations can replace `WordAdapter` and `MastraAdapter` without Office.js or a real backend. |
| **Extensibility** | Swapping the backend (e.g., from Mastra to a direct LLM API) requires only a new `IAnalysisPort` adapter. Swapping the document host (e.g., Google Docs) requires only a new `IDocumentPort` adapter. |
| **Maintainability** | `wordApi.ts` (579 lines, 4 responsibilities) is replaced by three focused files: `WordAdapter.ts`, `OoxmlPackageBuilder.ts`, `CommentCleanup.ts`. |
| **Scalability** | New analysis phases are added as a new `PipelineHandler` in the chain. No modifications to existing handlers or the orchestrator. |

---

## Module Map

```
src/
├── domain/                         ← Zero framework dependencies
│   ├── types.ts                    ← Shared interfaces (no runtime code)
│   ├── ports.ts                    ← IDocumentPort, IAnalysisPort
│   ├── review/
│   │   ├── DocumentReviewStateMachine.ts ← explicit document-review UI semantics
│   │   └── ReviewSessionMediator.ts      ← coordinates resolution + taskpane review state
│   └── pipeline/
│       ├── PipelineContext.ts      ← Shared state between handlers
│       ├── PipelineStateMachine.ts ← State machine (State pattern)
│       ├── PipelineEvents.ts       ← Event emitter (Observer pattern)
│       ├── PipelineOrchestrator.ts ← Chain of Responsibility runner
│       └── handlers/
│           ├── ReadTextHandler.ts
│           ├── CheckConnectionHandler.ts
│           ├── ChunkTextHandler.ts
│           ├── AnalyzeChunksHandler.ts
│           ├── DeduplicateHandler.ts
│           ├── GuardAppliedHandler.ts
│           └── ApplySuggestionsHandler.ts
│
├── adapters/
│   ├── word/
│   │   ├── WordAdapter.ts          ← implements IDocumentPort
│   │   ├── ApplySuggestionCommand.ts ← Command pattern
│   │   ├── ooxml/
│   │   │   └── OoxmlPackageBuilder.ts ← Builder pattern
│   │   └── cleanup/
│   │       └── CommentCleanup.ts   ← Range Colocation pattern
│   ├── mastra/
│   │   └── MastraAdapter.ts        ← implements IAnalysisPort
│   └── RetryAnalysisDecorator.ts   ← Decorator pattern
│
├── infrastructure/
│   ├── config.ts                   ← Constants (URLs, retry policy)
│   └── chunker.ts                  ← Pure text splitting function
│
└── taskpane/
    ├── taskpane.ts                 ← UI only: events, observers, render
    ├── taskpane.html
    └── taskpane.css
```

---

## Design Patterns

### GoF Patterns — Currently Used

| Pattern | Category | Location | Description |
|---------|----------|----------|-------------|
| **Chain of Responsibility** | Behavioral | `domain/pipeline/handlers/` | Each of the 7 analysis phases is an independent handler. The orchestrator runs them in sequence; any handler can abort the chain by setting `ctx.aborted = true`. |
| **Command** | Behavioral | `adapters/word/ApplySuggestionCommand.ts` | Each suggestion is encapsulated as a `DocumentCommand` with an `execute()` method. Enables future `undo()` support without restructuring. |
| **Observer** | Behavioral | `domain/pipeline/PipelineEvents.ts` | `PipelineEventEmitter` notifies registered `PipelineObserver` instances of phase starts, progress, completion, and aborts. The UI registers one observer; future analytics can register another without touching the pipeline. |
| **Mediator** | Behavioral | `domain/review/ReviewSessionMediator.ts` | Centralizes coordination between resolution workflow, document review state machine, cleanup visibility, and taskpane-facing review semantics. Prevents `taskpane.ts` from orchestrating cross-layer policy manually. |
| **State** | Behavioral | `domain/pipeline/PipelineStateMachine.ts` | Explicit state transitions (`idle → reading → connecting → chunking → analyzing → applying → done/error`) prevent concurrent runs and make lifecycle visible in code. |
| **State** | Behavioral | `domain/review/DocumentReviewStateMachine.ts` | Translates the authoritative `DocumentReviewState` snapshot into explicit document-review UI states (`idle`, `pending-review`, `ready-to-disable-track-changes`) so Track Changes CTA rules stop living in scattered booleans. |
| **Strategy** | Behavioral | `adapters/word/ApplySuggestionCommand.ts` | `classifyChange()` selects `insert`, `delete`, or `replace` tracked-change type based on suggestion content. |
| **Template Method** | Behavioral | `adapters/word/ApplySuggestionCommand.ts` | `execute()` has a fixed algorithm skeleton: search → extract format → disable tracking → build OOXML → insert → restore tracking. |
| **Iterator** | Behavioral | `handlers/AnalyzeChunksHandler.ts`, `adapters/word/WordAdapter.ts` | Sequential iteration over chunks and suggestions via `for...of`. |
| **Singleton** | Creational | `adapters/mastra/MastraAdapter.ts` | Single `MastraClient` instance reused across all workflow calls. |
| **Builder** | Creational | `adapters/word/ooxml/OoxmlPackageBuilder.ts` | Fluent API for constructing flat OPC OOXML packages: `.withDeletion().withInsertion().withComment().build()`. Replaces 120-line string concatenation. |
| **Factory Method** | Creational | `adapters/word/ApplySuggestionCommand.ts` | `classifyChange()` acts as a factory for `ChangeType` values used in OOXML construction. |
| **Facade** | Structural | `adapters/word/WordAdapter.ts` | Exposes a clean `IDocumentPort` interface hiding the complexity of `Word.run`, `context.sync`, and OOXML package structure. |
| **Decorator** | Structural | `adapters/RetryAnalysisDecorator.ts` | Wraps `IAnalysisPort` transparently to add retry-with-exponential-backoff without modifying `MastraAdapter`. |

### Ports & Adapters (Hexagonal Architecture)

| Port | Interface | Adapters |
|------|-----------|---------|
| Document Port | `IDocumentPort` | `WordAdapter` (Office.js) |
| Analysis Port | `IAnalysisPort` | `RetryAnalysisDecorator` → `MastraAdapter` (@mastra/client-js) |

### Important ownership note: Track Changes lifecycle

The historical implementation uses a preserve-and-restore pattern inside `ApplySuggestionCommand.ts` to toggle `changeTrackingMode` per suggestion. That explains the current behavior, but it should **not** be treated as the desired long-term architectural ownership model.

The corrected target direction is:

- per-suggestion commands mutate one suggestion,
- workflow/application layers own Track Changes lifecycle policy,
- `ReviewSessionMediator` is the explicit coordinator for taskpane-facing review policy,
- document-derived pending review state determines whether Track Changes should remain enabled,
- `DocumentReviewStateMachine` is the explicit application model that converts that document snapshot into taskpane-visible review states,
- the UI offers a final CTA when zero pending Stylistic artifacts remain.

See [`review-domain-and-track-changes.md`](./review-domain-and-track-changes.md) for the full requirement decisions and future-aligned architecture.

### Important ownership note: replace suggestion identity

The project has now identified a deeper architectural risk behind recurrent
`already-resolved` regressions in real Word:

- a native replace suggestion is semantically composed of an inserted side and a
  deleted side,
- but the current operational handle (`ContentControl` on `insertedRange`) is
  only attached to one side,
- and later resolution tries to reconstruct the whole replace by spatial
  inference.

That means the current code path must **not** be interpreted as a reliable long-
term identity model.

The corrected architectural direction is:

- one replace suggestion = one domain identity,
- multiple Word artifacts may reference that identity,
- lack of observable tracked changes must not be upgraded to confirmed
  `already-resolved`.

See [`replace-suggestion-identity-proposal.md`](./replace-suggestion-identity-proposal.md)
for the detailed proposal.

### Domain Patterns (Non-GoF)

| Pattern | Location | Description |
|---------|----------|-------------|
| **Preserve-and-Restore** | `ApplySuggestionCommand.ts` | `changeTrackingMode` saved before modification, restored in `finally` even on error. |
| **Guard Clause** | `GuardAppliedHandler.ts` | Filters suggestions already applied as tracked changes. Prevents duplicates when user re-runs analysis. |
| **Partial Success** | `AnalyzeChunksHandler.ts`, `WordAdapter.ts` | Chunk failures and suggestion failures are collected, not fatal. User gets maximum possible value. |
| **Fail-Fast Gate** | `CheckConnectionHandler.ts` | Verifies backend connectivity before starting analysis. Aborts immediately if unavailable. |
| **Range Colocation** | `CommentCleanup.ts` | Uses `Range.compareLocationWith()` to detect orphaned comments. Document is the source of truth — no in-memory registry. |
| **Transparent Fallback** | `WordAdapter.getTextToAnalyze()` | Returns selection if active, full document otherwise. Caller receives `{ text, isSelection }` without knowing how it was resolved. |
| **Null Return on Error** | `MastraAdapter.analyzeChunk()` | Never throws; always returns `ChunkResult` (with empty suggestions on failure). Enables partial success. |
| **Retry + Exponential Backoff** | `RetryAnalysisDecorator.ts` | `delay = baseMs * 2^attempt` between retries. 3 max attempts. Separated from `MastraAdapter` via Decorator. |
| **Per-Resource Isolation** | `ApplySuggestionCommand.ts` | Each suggestion runs in its own `Word.run` context to avoid stale ranges after OOXML insertions shift document positions. |
| **Composition Root** | `taskpane/taskpane.ts` | Single wiring point: instantiates adapters, decorators, orchestrator, and state machine. No other module knows the full dependency graph. |
| **Derived Snapshot + Explicit UI State** | `WordAdapter.ts`, `domain/review/DocumentReviewStateMachine.ts`, `domain/review/ReviewSessionMediator.ts`, `taskpane.ts` | Word remains the source of truth through `DocumentReviewState`, the state machine centralizes review UI semantics, and the mediator coordinates cleanup/CTA/taskpane consequences as one policy surface. |
| **Compound Identity (proposed)** | `docs/replace-suggestion-identity-proposal.md` | A replace suggestion should be modeled as one domain identity with multiple Word artifact references, instead of treating one inserted-side `ContentControl` as the whole suggestion. |

> **Documentation note:** Preserve-and-Restore accurately describes the current code path, but the Track Changes requirement change intentionally moves the desired ownership model away from per-suggestion toggling. Keep that distinction explicit when editing this section in the future.

---

## Data Flow

### Analysis Pipeline (Chain of Responsibility)

```
User clicks "Analizar y sugerir"
        │
        ▼
taskpane.ts: handleAnalyze()
  - Creates PipelineContext { documentPort, analysisPort, emitter, profile }
  - Registers UI PipelineObserver on emitter
  - stateMachine.transition("reading")
        │
        ▼
PipelineOrchestrator.run(ctx)
        │
        ├──► ReadTextHandler
        │       └── documentPort.getTextToAnalyze() → ctx.text, ctx.isSelection
        │       └── Abort if text is empty
        │
        ├──► CheckConnectionHandler
        │       └── analysisPort.checkConnection() → fail-fast gate
        │       └── Abort if backend unreachable
        │
        ├──► ChunkTextHandler
        │       └── splitText(ctx.text, maxChunkSize) → ctx.chunks
        │
        ├──► AnalyzeChunksHandler
        │       └── For each chunk (sequential):
        │             analysisPort.analyzeChunk(chunk, profile, "es") → suggestions
        │       └── Collects ctx.rawSuggestions, ctx.chunkErrors
        │       └── Abort if zero suggestions
        │
        ├──► DeduplicateHandler
        │       └── Removes cross-chunk duplicates (case-insensitive)
        │       └── Sets ctx.uniqueSuggestions
        │
        ├──► GuardAppliedHandler
        │       └── documentPort.getAppliedOriginalTexts() → Set<string>
        │       └── Filters already-applied suggestions
        │       └── Sets ctx.pendingSuggestions
        │       └── Abort if nothing pending
        │
        └──► ApplySuggestionsHandler
                └── documentPort.applySuggestions(pending, onProgress)
                      └── For each suggestion: new ApplySuggestionCommand(s).execute()
                            └── Word.run: search → extract format → build OOXML → insert
                └── emitter.emitComplete(suggestions, result, errors, isSelection)
                └── Sets ctx.result

taskpane.ts: PipelineObserver.onComplete()
  └── renderResults() → updates DOM
  └── Rehydrates `DocumentReviewStateMachine` from `result.pendingAfter`
  └── Applies cleanup/CTA visibility from explicit review state

### Document Review State Flow

```
Word document (authoritative host state)
        │
        ├──► WordAdapter.inspectDocumentReviewState()
        │       └── derives DocumentReviewState {
        │             pendingStylisticArtifacts,
        │             hasPendingStylisticArtifacts,
        │             trackChangesActive
        │           }
        │
        ├──► DocumentReviewStateMachine.deriveState()
        │       └── maps snapshot to one explicit UI state:
        │             - idle
        │             - pending-review
        │             - ready-to-disable-track-changes
        │
        ├──► ReviewSessionMediator
        │       └── combines document review state + cleanup preview +
        │           resolution workflow result into one taskpane-facing state
        │
        └──► taskpane.ts
                └── renders from mediator output instead of cross-layer glue logic
```

### Replace Suggestion Identity Risk

```text
Logical replace suggestion
        │
        ├── inserted-side Word artifact
        ├── deleted-side Word artifact
        └── operational anchor(s)

Current risk:
taskpane/workflow may receive "already-resolved"
from an observation strategy that only proved
"no tracked changes were observed near the inserted-side CC".

That is not sufficient proof of semantic resolution.
```

### Why this state machine was added

Originally, document-review UI semantics were inferred in multiple places with
booleans like:

- `hasPendingStylisticArtifacts`
- `trackChangesActive`
- `showDisableTrackChangesCta`

That approach made the code vulnerable to UI inconsistency because the same rule
could be reinterpreted differently by the adapter, workflow, and taskpane.

The current correction keeps the architectural distinction explicit:

- `DocumentReviewState` = authoritative snapshot read from Word
- `DocumentReviewStateMachine` = explicit frontend state model derived from that snapshot
- `taskpane.ts` = renders from the state model, not from ad-hoc boolean combinations
```

### Comment Cleanup Flow (Direct port call, no pipeline)

```
User clicks "Limpiar comentarios resueltos"
        │
        ▼
taskpane.ts: handleCleanup()
        └── documentPort.cleanupResolvedComments()
              └── CommentCleanup.cleanupResolvedComments()
                    ├── Sync 1: Load Stylistic comments and tracked changes
                    ├── Sync 2: Get document ranges for each
                    ├── Sync 3: Compare every comment range vs every TC range
                    └── Sync 4: Delete orphaned comments, keep colocated ones
```

---

## OOXML Strategy

All changes are applied via flat OPC OOXML packages built by `OoxmlPackageBuilder`. Each package contains 4 parts:

1. **`/_rels/.rels`** — Package relationships (points to document.xml)
2. **`/word/_rels/document.xml.rels`** — Document relationships (points to comments.xml)
3. **`/word/document.xml`** — Tracked change markup (`<w:del>` + `<w:ins>`) with comment anchors
4. **`/word/comments.xml`** — Formatted justification (bold category + justification text)

Builder usage:
```typescript
const ooxml = new OoxmlPackageBuilder()
  .withRunProperties(runPropsXml)          // preserve original formatting
  .withChange(original, replacement, type, "Stylistic", isoDate)
  .withComment(category, justification, "Stylistic", isoDate)
  .build();
```

This approach ensures:
- Tracked change blue card shows `w:author="Stylistic"` cleanly
- Comment appears as a margin balloon with bold category + justification
- Original text formatting (`<w:rPr>`) is preserved
- No double-tracking: `changeTrackingMode` is disabled during insertion (Preserve-and-Restore)

---

## Four Pillars Evaluation

| Pillar | Status | Evidence |
|--------|--------|---------|
| **Maintainability** | ✅ Good | Each module has one responsibility. `wordApi.ts` (579 lines, 4 concerns) replaced by 3 focused files. `taskpane.ts` now only handles DOM events and rendering. |
| **Testability** | ✅ Good | Pipeline depends only on `IDocumentPort` and `IAnalysisPort` interfaces. Mock adapters enable unit tests with no Office.js or Mastra dependency. `chunker.ts` and `deduplicateByOriginalText` (inside `DeduplicateHandler`) are pure functions. |
| **Scalability** | ✅ Good | Chunking handles 200K+ word documents. Partial success philosophy maximizes results on large documents. New analysis phases are added as a new handler — O(1) change regardless of pipeline size. |
| **Extensibility** | ✅ Good | New backend → new `IAnalysisPort` adapter (zero pipeline changes). New document host → new `IDocumentPort` adapter (zero pipeline changes). New analysis phase → new `PipelineHandler` inserted into orchestrator array. New cross-cutting concern → new `Decorator` wrapping a port. |

---

## Key Design Decisions

### Why Chain of Responsibility for the pipeline?

`handleAnalyze()` was a monolithic 160-line function with 7 phases. Chain of Responsibility makes each phase an independent, testable, replaceable handler. Adding a new phase (e.g., terminology verification, spellcheck pre-filter) requires creating one new file and inserting it into the handler array — no modifications to existing handlers or the orchestrator.

### Why Command for suggestion application?

Each suggestion is encapsulated as a `DocumentCommand` with `execute()`. This enables:
- Clear single responsibility: one command = one document mutation
- Future `undo()` support without restructuring
- Per-command result tracking (`CommandResult`)

### Why Decorator for retry?

The original `mastraClient.ts` mixed retry logic with Mastra communication. The Decorator separates concerns: `MastraAdapter` only handles the HTTP protocol; `RetryAnalysisDecorator` only handles retry semantics. Tests can inject the bare adapter without retry overhead. Future concerns (circuit breaker, caching) can be added as additional decorators.

### Why Observer for progress?

The `ProgressCallback` was a single listener. `PipelineEventEmitter` supports multiple simultaneous observers. The UI subscribes one observer; future analytics or structured logging can subscribe additional observers without touching the pipeline.

### Why State Machine for pipeline lifecycle?

The `isRunning` boolean check (`btn.disabled`) was implicit state. `PipelineStateMachine` makes valid transitions explicit and validated. Concurrent pipeline runs are prevented at the state level — impossible to go from `idle` to `applying` without passing through all intermediate states.

### Why one Word.run per suggestion?

After an OOXML insertion replaces a range, all subsequent ranges in the document may shift. A fresh `Word.run` per suggestion means each search starts from a clean document state. A failure in suggestion N doesn't affect suggestions 1..(N-1). This is the Per-Resource Isolation pattern.

### Why Builder for OOXML?

`buildTrackedChangeOoxml()` was 120 lines of string concatenation, opaque and hard to extend. `OoxmlPackageBuilder` provides a fluent API that makes intent clear and extension easy: adding a new OOXML part (e.g., footnotes, format-only changes) is additive, not invasive.

### Why Range Colocation for comment cleanup?

Comments and tracked changes have no direct link in the Word API. Using `Comment.getRange()` and `TrackedChange.getRange()` with `Range.compareLocationWith()` to determine colocation:
- Uses the document as the source of truth (no in-memory registry)
- Works across sessions (no state to persist)
- Never deletes comments it can't positively identify as orphaned

---

## Error Handling

Errors are handled at four levels:

| Level | Location | Strategy |
|-------|----------|---------|
| **Word API** | `ApplySuggestionCommand.ts` | `try/finally` ensures `changeTrackingMode` is always restored. Each command is independent. |
| **Backend** | `RetryAnalysisDecorator.ts` | Retry with exponential backoff (3 attempts). Never throws — returns empty result. |
| **Pipeline** | `handlers/` | Handlers set `ctx.aborted = true` for graceful abort. `emitter.emitAbort(reason)` notifies observers. |
| **UI** | `taskpane.ts` | `try/catch/finally` around `orchestrator.run()`. `setAnalyzeLoading(false)` always runs in `finally`. |

### Partial Success Philosophy

The system is designed for partial success:
- If 3/5 chunks succeed → suggestions from those 3 are applied
- If 25/30 suggestions are found → 25 are applied, 5 reported as "not found"
- Results panel always shows the complete picture: applied count, failed count, chunk errors
