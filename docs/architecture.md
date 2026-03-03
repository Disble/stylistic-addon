# Architecture

This document describes the layered architecture of Stylistic, the data flow between modules, and the key design decisions behind them.

## Design Principles

1. **Separation of concerns** — each module owns a single responsibility. Business logic never touches Office.js. The UI never calls Word directly.
2. **Preserve-and-restore** — the add-in always reads the document's `changeTrackingMode` before modifying it and restores it when done, even if an error occurs (via `try/finally`).
3. **Batch operations** — all `body.search()` calls are enqueued before a single `context.sync()`, minimizing round-trips to the Word host.
4. **Pure analyzers** — the analysis layer (`analyzer.ts`) has zero dependencies on Office.js or the DOM, making it testable with any JavaScript test runner.

## Module Map

```
src/
├── lib/
│   ├── types.ts        ← Shared interfaces (no runtime code)
│   ├── wordApi.ts      ← Word API abstraction (only Office.js consumer)
│   └── analyzer.ts     ← Text analysis engine (pure business logic)
└── taskpane/
    ├── taskpane.ts      ← UI orchestrator (event binding, rendering)
    ├── taskpane.html    ← Task pane markup
    └── taskpane.css     ← Styling
```

### Dependency Graph

```
taskpane.ts
├── imports wordApi.ts
├── imports analyzer.ts
└── imports types.ts

wordApi.ts
├── imports types.ts
└── uses global: Word (Office.js)

analyzer.ts
└── imports types.ts

types.ts
└── (no imports)
```

The dependency graph is strictly acyclic. `types.ts` is at the bottom. `wordApi.ts` and `analyzer.ts` are peers that never import each other. `taskpane.ts` is the composition root.

## Data Flow

The main analysis flow follows this sequence:

```
User clicks "Analizar y sugerir"
        │
        ▼
taskpane.ts: handleAnalyze()
        │
        ├──► wordApi.getDocumentText()
        │         │
        │         └── Word.run → body.load("text") → context.sync()
        │         └── returns plain text string
        │
        ├──► analyzer.analyze(text)
        │         │
        │         ├── redundancyRule.detect(text)
        │         ├── fillerRule.detect(text)
        │         ├── wordChoiceRule.detect(text)
        │         └── deduplicateByOriginalText()
        │         └── returns Suggestion[]
        │
        ├──► wordApi.insertSuggestionsAsTrackedChanges(suggestions)
        │         │
        │         ├── Load & save current changeTrackingMode
        │         ├── Set changeTrackingMode = TrackAll
        │         ├── Batch: body.search() for each suggestion
        │         ├── context.sync() (single round-trip)
        │         ├── Replace matches with insertText()
        │         ├── context.sync()
        │         └── Restore previous changeTrackingMode (in finally block)
        │         └── returns InsertionResult
        │
        └──► taskpane.ts: renderResults(suggestions, result)
                  └── Displays applied/failed suggestions in the task pane
```

## Key Design Decisions

### Why vanilla HTML instead of React?

The task pane UI is minimal: two buttons, a preview area, a results list, and a status bar. React would add ~40KB to the bundle and introduce a build-time dependency for no measurable benefit. If the UI grows significantly, React can be added incrementally — the architecture doesn't prevent it.

### Why `try/finally` in `insertSuggestionsAsTrackedChanges`?

If an error occurs mid-insertion (e.g., a search fails, the document becomes read-only), the tracking mode must still be restored to its original value. Without the `finally` block, a crash could leave the document in `TrackAll` mode permanently, which is surprising and disruptive for the user.

### Why batch search before sync?

Each `context.sync()` is a round-trip to the Word host process. Calling `body.search()` inside a loop with a sync per iteration would be O(n) round-trips. By enqueuing all searches and syncing once, we reduce it to O(1) round-trips regardless of the number of suggestions.

### Why deduplicate suggestions?

Multiple rules can flag the same text. For example, a phrase might be both a filler and a redundancy. Without deduplication, the second replacement would fail (the original text is already replaced) and appear as a false "not found" error. Deduplication by `originalText` prevents this.

### Why case-sensitive search with case-insensitive deduplication?

The `body.search()` uses `matchCase: true` to avoid false positives (e.g., matching a proper noun). But deduplication uses case-insensitive comparison because "Utilizar" and "utilizar" are the same word for editorial purposes.

## Error Handling Strategy

Errors are handled at three levels:

| Level | Module | Strategy |
|---|---|---|
| Word API | `wordApi.ts` | `try/finally` ensures tracking mode is always restored. Errors propagate to the caller. |
| Orchestrator | `taskpane.ts` | `try/catch` around every handler. Errors are translated to user-friendly messages via `toUserMessage()`. |
| UI | `taskpane.ts` | `setAnalyzeLoading(false)` runs in `finally` to ensure the button is always re-enabled. |

Known Office.js error codes are mapped to Spanish messages:

| Error Code | User Message |
|---|---|
| `AccessDenied` | "El documento está protegido o es de solo lectura." |
| `InvalidArgument` | "Argumento inválido al comunicarse con Word." |
| `ItemNotFound` | "No se encontró el elemento solicitado en el documento." |
| (other) | Raw error message from Office.js |
