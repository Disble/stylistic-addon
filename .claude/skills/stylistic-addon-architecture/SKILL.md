---
name: stylistic-addon-architecture
description: >
  Architecture guide for the stylistic-addon Word Add-in project.
  Trigger: When writing any code for this project, adding new features,
  creating adapters, ports, or modifying taskpane.ts. Load before writing any code.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "1.0"
---

# Stylistic Add-in — Architecture Skill

Load this skill BEFORE writing ANY code in this project. It is the authoritative reference for how this codebase is structured, what patterns are required, and what is strictly forbidden.

---

## 1. Project Identity

- **Type:** Microsoft Word Office Add-in (Task Pane)
- **Language:** TypeScript 5.4+
- **Bundler:** Webpack 5
- **UI:** Pure DOM — NO UI framework. No React, Vue, or Angular. Ever.
- **Backend:** Mastra AI framework (`@mastra/client-js` v1.7.1)
- **Test runner:** Vitest (globals: true, environment: node)
- **Dev server port:** 3000 | **Mastra default port:** 4111

---

## 2. Hexagonal Architecture — Strict

The dependency rule is **inviolable**: outer layers depend on inner layers. Inner layers NEVER import from outer layers.

### Frontend domain definition (authoritative)

When reasoning about this project, do **not** define the addon as “just a client for the backend” and do **not** confuse internal review state with the full domain.

The frontend domain is:

> **Presentar al usuario sugerencias de estilo provenientes del backend y permitirle aceptarlas, rechazarlas o debatirlas dentro de Word.**

Use this definition to evaluate names, layering, workflow ownership, and future refactors.

### Minimum ubiquitous language

- **`ReviewSuggestion`** — a suggestion already contextualized in Word, with frontend-owned state and interactions.
- **`DocumentReviewState`** — the logical state of Stylistic review artifacts as they exist in the document.
- **`ReviewProcessState`** — temporary workflow/process state (`reading`, `analyzing`, `applying`, etc.).
- **`Debate`** — future first-class capability; not equivalent to current fire-and-forget feedback.

If a change blurs those boundaries, stop and correct the model before writing code.

```
┌──────────────────────────────────────────────────────────────────┐
│  PRESENTATION — src/taskpane/                                    │
│  taskpane.ts — event binding, observer registration, rendering   │
│  Composition Root: instantiates all adapters and wires the graph │
├──────────────────────────────────────────────────────────────────┤
│  DOMAIN — src/domain/                                            │
│  types.ts     — pure TypeScript interfaces (zero runtime code)  │
│  ports.ts     — IDocumentPort, IAnalysisPort, IFeedbackPort      │
│  pipeline/    — Chain of Responsibility: handlers + orchestrator │
│                 State machine, event emitter, pipeline context   │
├───────────────────────────┬──────────────────────────────────────┤
│  ADAPTER: Word            │  ADAPTER: Mastra + Feedback          │
│  src/adapters/word/       │  src/adapters/mastra/                │
│  WordAdapter.ts           │  MastraAdapter.ts                    │
│  (implements IDocumentPort│  FeedbackAdapter.ts                  │
│   via Office.js)          │  MockFeedbackAdapter.ts              │
│                           │  (implements IAnalysisPort,          │
│                           │   IFeedbackPort)                     │
├───────────────────────────┴──────────────────────────────────────┤
│  INFRASTRUCTURE — src/infrastructure/                            │
│  config.ts   — ALL constants (URLs, IDs, limits, retry policy)  │
│  chunker.ts  — pure text splitting function (no side effects)   │
└──────────────────────────────────────────────────────────────────┘
```

### Full Module Map

```
src/
├── domain/                         ← Zero framework dependencies
│   ├── types.ts                    ← Shared interfaces (no runtime code)
│   ├── ports.ts                    ← IDocumentPort, IAnalysisPort, IFeedbackPort
│   └── pipeline/
│       ├── PipelineContext.ts      ← Shared mutable state between handlers
│       ├── PipelineStateMachine.ts ← State pattern (idle→reading→…→done/error)
│       ├── PipelineEvents.ts       ← Observer pattern (PipelineEventEmitter)
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
│   │   ├── WordAdapter.ts              ← Facade: implements IDocumentPort
│   │   ├── ApplySuggestionCommand.ts   ← Command + Strategy + Template Method
│   │   ├── ooxml/
│   │   │   └── OoxmlPackageBuilder.ts  ← Builder pattern (fluent API)
│   │   └── cleanup/
│   │       └── CommentCleanup.ts       ← Range Colocation pattern
│   ├── mastra/
│   │   ├── MastraAdapter.ts            ← implements IAnalysisPort (Singleton client)
│   │   ├── FeedbackAdapter.ts          ← implements IFeedbackPort (production)
│   │   └── MockFeedbackAdapter.ts      ← implements IFeedbackPort (dev stub)
│   └── RetryAnalysisDecorator.ts       ← Decorator: wraps IAnalysisPort
│
├── infrastructure/
│   ├── config.ts                   ← ALL constants live here
│   └── chunker.ts                  ← Pure function: splitText()
│
└── taskpane/
    ├── taskpane.ts                 ← Composition Root + UI only
    ├── taskpane.html
    └── taskpane.css
```

---

## 3. Ports and Adapters Table

| Port Interface | Production Adapter | Dev/Test Adapter | Location |
|---|---|---|---|
| `IDocumentPort` | `WordAdapter` (Office.js) | `MockWordAdapter` (if needed) | `adapters/word/` |
| `IAnalysisPort` | `RetryAnalysisDecorator` → `MastraAdapter` | `MockAnalysisAdapter` (if needed) | `adapters/mastra/`, `adapters/` |
| `IFeedbackPort` | `FeedbackAdapter` | `MockFeedbackAdapter` | `adapters/mastra/` |

**The `RetryAnalysisDecorator` wraps `IAnalysisPort` transparently.** It adds retry-with-exponential-backoff without touching `MastraAdapter`. It is instantiated in `taskpane.ts` as:
```ts
const analysisPort = new RetryAnalysisDecorator(
  new MastraAdapter(),
  MAX_RETRIES,
  RETRY_BASE_DELAY_MS
);
```

---

## 4. Backend Communication — `@mastra/client-js` ONLY

**NEVER use raw `fetch` for backend calls.** The ONLY backend communication pattern in this codebase is:

### Submit + Poll pattern (IAnalysisPort — editorial workflow)

```ts
// Submit a chunk
const workflow = client.getWorkflow(WORKFLOW_ID);
const run = await workflow.createRun();
// run.runId is extracted here for later polling
await run.start({ inputData: payload });

// Poll until terminal
const state = await workflow.runById(runId, {
  fields: ["result", "error"],   // NEVER add "status" here — causes HTTP 400
  withNestedWorkflows: false,
});
```

### Fire-and-forget pattern (IFeedbackPort — feedback workflow)

```ts
const workflow = mastraClient.getWorkflow(FEEDBACK_WORKFLOW_ID);
const run = await workflow.createRun();
await run.start({ inputData: payload });
// No polling — fire-and-forget
```

**SDK notes for `@mastra/client-js` v1.7.1:**
- `.execute()` does NOT exist — use `createRun()` + `run.start()`
- `status` is ALWAYS in the response metadata — do NOT include it in `fields` (causes HTTP 400)
- `runId` is extracted from the `createRun()` return value before calling `run.start()`

---

## 5. Adding a New Port / Adapter — Step by Step

1. **Add the payload type** to `src/domain/types.ts` (pure interface, no imports from frameworks)
2. **Add the port interface** to `src/domain/ports.ts` (import only from `./types`)
3. **Add config constants** to `src/infrastructure/config.ts` (workflow IDs, URLs, limits)
4. **Create the production adapter** in `src/adapters/{system}/{AdapterName}.ts`
   - Implements the port interface
   - Imports `MastraClient` or Office.js — NEVER from domain
5. **Create a mock adapter** in the same directory (`Mock{AdapterName}.ts`)
   - Logs to console, no network calls
   - Used in dev until backend is ready
6. **Wire it in `taskpane.ts`** (Composition Root — module level):
   ```ts
   const myPort: IMyPort = new MockMyAdapter(); // swap to MyAdapter in prod
   ```
7. **Write tests** using the mock adapter (never the real adapter in tests)

---

## 6. DOM Rendering — Programmatic Only

All UI rendering is done via programmatic DOM in `taskpane.ts`. The render functions (e.g., `renderResults()`) follow these rules:

- **ALWAYS** use `document.createElement()` + `.appendChild()` / `.textContent`
- **NEVER** use `innerHTML` with dynamic/user-derived data (XSS risk)
- **No UI framework** — no JSX, no templates, no virtual DOM
- CSS max-height transitions are used for accordion animations (no JS animation)

```ts
// Correct
const li = document.createElement("li");
li.textContent = suggestion.category;   // safe, no XSS
container.appendChild(li);

// FORBIDDEN
container.innerHTML = `<li>${suggestion.category}</li>`;  // NEVER with dynamic data
```

---

## 7. Fire-and-Forget Pattern

When invoking `IFeedbackPort.sendFeedback()`, use `void` to make the intent explicit:

```ts
void feedbackPort.sendFeedback(payload);   // intentionally not awaited
```

Rules:
- Do NOT `await` feedback calls in the UI
- Feedback adapters MUST wrap their bodies in `try/catch` and swallow all errors silently
- The user must never see a feedback transmission failure
- No retry logic for feedback

```ts
// FeedbackAdapter (production) — correct implementation
async sendFeedback(payload: FeedbackPayload): Promise<void> {
  try {
    const workflow = mastraClient.getWorkflow(FEEDBACK_WORKFLOW_ID);
    const run = await workflow.createRun();
    await run.start({ inputData: payload });
  } catch {
    // Swallow all errors silently — fire-and-forget
  }
}
```

---

## 8. TDD Policy

- **Test runner:** Vitest (`npm test` / `vitest run`)
- **Config:** `vitest.config.ts` — globals: true, environment: node
- **Pattern:** Vertical slices — RED → GREEN → REFACTOR per behavior
- **NEVER** write horizontal tests (covering multiple behaviors in one test)
- Tests verify behavior through **public interfaces only** — never test private methods
- Mock adapters replace all external dependencies (Office.js, Mastra backend)
- Test files: `src/**/*.test.ts` (colocated with source files)
- Excluded from tests: `src/commands/**` (Office commands infrastructure)

```ts
// Vertical slice — one behavior per test
it("returns empty suggestions when workflow returns no matches", async () => {
  const adapter = new MockAnalysisAdapter();
  // ... test one specific behavior
});
```

---

## 9. Config Pattern — All Constants in `src/infrastructure/config.ts`

**EVERY** magic value must be a named export from `config.ts`. Never hardcode values elsewhere.

```ts
// src/infrastructure/config.ts — current constants
export const MASTRA_BASE_URL = "http://localhost:4111";
export const WORKFLOW_ID = "stylistic-workflow";
export const FEEDBACK_WORKFLOW_ID = "feedback-workflow";
export const DEFAULT_MAX_CHUNK_SIZE = 100_000;
export const WORD_API_BATCH_SIZE = 30;
export const MAX_RETRIES = 3;
export const RETRY_BASE_DELAY_MS = 1_000;
export const POLL_INTERVAL_MS = 1_000;
export const DEFAULT_PROFILES: Profile[] = [ /* ... */ ];
```

---

## 10. Mock Swap Strategy — Composition Root

`taskpane.ts` is the **only** Composition Root. All adapters are instantiated at module level:

```ts
// src/taskpane/taskpane.ts — module-level instantiation (current state)
const documentPort = new WordAdapter();
const analysisPort = new RetryAnalysisDecorator(
  new MastraAdapter(),
  MAX_RETRIES,
  RETRY_BASE_DELAY_MS
);
const feedbackPort = new MockFeedbackAdapter(); // ← swap to FeedbackAdapter when backend is ready
```

Switching from mock to production is **always a one-line change** in `taskpane.ts`. No other file changes.

---

## 11. Pipeline — Chain of Responsibility

The 7-handler pipeline is orchestrated by `PipelineOrchestrator`. Each handler:
- Receives a `PipelineContext` (shared mutable state)
- Can read/write `ctx` fields
- Can abort the chain by setting `ctx.aborted = true`
- Emits events via `ctx.emitter` (Observer pattern)

**Adding a new analysis phase:**
1. Create `src/domain/pipeline/handlers/MyNewHandler.ts`
2. Insert it into the array in `taskpane.ts` `PipelineOrchestrator([..., new MyNewHandler(), ...])`
3. No changes to `PipelineOrchestrator` or other handlers

**Pipeline state transitions** (enforced by `PipelineStateMachine`):
```
idle → reading → connecting → chunking → analyzing → applying → done
                                                                → error
```

### Architectural direction for suggestion resolution

The analysis pipeline is already Chain of Responsibility. Future work that touches accept/reject resolution must respect the same architectural language.

Treat accept/reject as a workflow, not as ad-hoc procedural glue between `taskpane.ts` and `WordAdapter`.

Required direction:

- resolution flow → **Chain of Responsibility**
- per-action variation (`accept` / `reject`) → **Strategy**
- feedback → included in the workflow result, but **never blocking** for UX

Do **not** let `taskpane.ts` become the long-term owner of review workflow semantics.

---

## 12. WorkflowInput Shape — Current Backend Contract

```ts
// src/domain/types.ts — WorkflowInput (sent to WORKFLOW_ID)
interface WorkflowInput {
  text: string;
  genero: "narrativa-literaria" | "ensayo-academico" | "periodismo-cultural" | "general";
  autorSlug: string;  // kebab-case author identifier
}
```

⚠️ **The backend uses `genero`/`autorSlug` (NOT `profile`/`language`)**. The `api-contract.md` docs may be out of date — the source of truth is `src/domain/types.ts`.

---

## 13. Critical Rules — NEVER DO

| Rule | Rationale |
|---|---|
| **NEVER** use raw `fetch` for backend calls | Only `@mastra/client-js` SDK allowed |
| **NEVER** import from frameworks/libraries into `src/domain/` | Domain must be pure TS, framework-free |
| **NEVER** put business logic in `taskpane.ts` | Composition Root is UI + wiring only |
| **NEVER** add a UI framework (React, Vue, Angular) | Pure DOM is the rule; add-ins don't need it |
| **NEVER** use `innerHTML` with dynamic/user data | XSS risk; use `textContent` + `createElement` |
| **NEVER** add `"status"` to `fields` in `workflow.runById()` | Causes deterministic HTTP 400 in SDK v1.7.1 |
| **NEVER** call `.execute()` on a workflow | Method does not exist in SDK v1.7.1 |
| **NEVER** hardcode URLs, IDs, or limits outside `config.ts` | Single source of truth for all constants |
| **NEVER** await `feedbackPort.sendFeedback()` in the UI | Fire-and-forget by design |
| **NEVER** let feedback errors surface to the user | Adapters must swallow all errors silently |

---

## 14. Design Patterns Reference

| Pattern | Location | Purpose |
|---|---|---|
| Chain of Responsibility | `domain/pipeline/handlers/` | 7 sequential analysis phases |
| Command | `adapters/word/ApplySuggestionCommand.ts` | Each suggestion = one encapsulated mutation |
| Observer | `domain/pipeline/PipelineEvents.ts` | UI and future analytics subscribe independently |
| State | `domain/pipeline/PipelineStateMachine.ts` | Prevent concurrent runs, explicit transitions |
| Strategy | `adapters/word/ApplySuggestionCommand.ts` | Selects insert/delete/replace OOXML type |
| Singleton | `adapters/mastra/MastraAdapter.ts` | Single `MastraClient` instance reused |
| Builder | `adapters/word/ooxml/OoxmlPackageBuilder.ts` | Fluent OOXML package construction |
| Decorator | `adapters/RetryAnalysisDecorator.ts` | Adds retry without modifying `MastraAdapter` |
| Facade | `adapters/word/WordAdapter.ts` | Hides Office.js complexity behind `IDocumentPort` |
| Ports & Adapters | Entire architecture | Testability and extensibility |

### Track Changes lifecycle rule (requirement change)

If you touch the Track Changes lifecycle, follow these rules:

1. **Never** toggle `changeTrackingMode` per suggestion as a desired target architecture.
2. Enable Track Changes **lazily**, only when the first real `track-change` suggestion is applied.
3. Keep Track Changes enabled while **any Stylistic artifact remains pending in Word**.
4. The source of truth for “pending” is the **document**, not the taskpane.
5. When pending Stylistic artifacts reach zero, do **not** auto-disable Track Changes.
6. Instead, expose a UI CTA with a single action: **`Desactivar control de cambios`**.

The current code may not fully implement these rules yet. Treat this skill as the architectural target.

---

## 16. Linting and Naming Enforcement

- Use **Biome** as the single formatter and linter for this add-in. It replaces `office-addin-lint`, `eslint-plugin-office-addins`, and Prettier in one tool.
- Enable Biome filename enforcement via `useFilenamingConvention`. Allowed cases: `PascalCase` (class files) and `camelCase` (module/utility files).
- Biome handles general filename case, but the project architecture enforces additional structural rules via `scripts/check-file-naming.mjs`:

### Structural naming rules (enforced in `src/domain/`, `src/adapters/`, `src/infrastructure/`)

| Rule | Description | Example violation |
|---|---|---|
| **No `utils.ts`/`Utils.ts`** | Generic utils name is an architecture smell; use a descriptive module name | `src/domain/utils.ts` |
| **Handler suffix** | All files in any `handlers/` directory must end with `Handler.ts` | `TextProcessor.ts` inside `handlers/` |
| **Adapter OOP suffix** | All non-test files in `src/adapters/` must end with a known suffix (after stripping `Mock` prefix) | `BackendClient.ts` in `adapters/` |
| **No triple-compound names** | At most one dot separator after the base name | `foo.helpers.types.ts` |

### Known adapter suffixes

```text
Adapter · Decorator · Command · Builder · Cleanup · Machine · Events · Context · Orchestrator
```

A `Mock` prefix is allowed in front of any suffix (e.g. `MockFeedbackAdapter.ts`).

### What each naming pattern signals

| Pattern | Meaning | Example |
|---|---|---|
| `*Adapter.ts` | Implements a port interface; hides framework/IO details | `WordAdapter.ts`, `MastraAdapter.ts` |
| `*Decorator.ts` | Wraps a port to add cross-cutting behavior transparently | `RetryAnalysisDecorator.ts` |
| `*Command.ts` | Encapsulates one document mutation; enables future undo | `ApplySuggestionCommand.ts` |
| `*Builder.ts` | Fluent construction API for a complex object | `OoxmlPackageBuilder.ts` |
| `*Cleanup.ts` | Range-colocation or document-cleanup operation | `CommentCleanup.ts` |
| `*Handler.ts` | One phase in the Chain of Responsibility pipeline | `ReadTextHandler.ts` |
| `Mock*.ts` | Test double; same interface, no side effects | `MockFeedbackAdapter.ts` |

---

## 17. Git Hook Workflow

- Use **Lefthook** for repository hooks; wiring is in `lefthook.yml`.
- Hooks are installed automatically on `npm install` via the `prepare` lifecycle script.
- `pre-commit` runs two fast checks only:
  1. **`lint:staged`** — Biome checks and auto-fixes staged TypeScript/JavaScript files.
  2. **`check:filenames`** — validates structural naming conventions across managed folders.
- `pre-push` runs one slow check:
  1. **`typecheck`** — full `tsc --noEmit` across the project. Catches type errors before they reach CI.
- Do NOT add builds or tests to `pre-commit`; slow checks belong in `pre-push`.
- If Biome auto-fixes a staged file (`stage_fixed: true`), Lefthook re-stages the fix automatically.

### Available scripts

```bash
npm run lint              # biome check . (full project)
npm run lint:write        # biome check --write . (full project, auto-fix)
npm run lint:staged       # biome check --staged --write (pre-commit safe)
npm run check:filenames   # structural naming validation only
npm run validate          # lint + check:filenames (full pre-push gate)
npm run hooks:install     # re-install lefthook hooks manually
npm run hooks:pre-commit  # run pre-commit hook locally without committing
```

---

## 15. Error Handling Layers

| Level | Location | Strategy |
|---|---|---|
| Word API | `ApplySuggestionCommand.ts` | `try/finally` — always restores `changeTrackingMode` |
| Backend transport | `RetryAnalysisDecorator.ts` | Retry + exponential backoff, 3 max attempts |
| Pipeline handlers | `handlers/` | `ctx.aborted = true` + `emitter.emitAbort(reason)` |
| UI | `taskpane.ts` | `try/catch/finally` — `setAnalyzeLoading(false)` always runs |
| Feedback | `FeedbackAdapter.ts` | `try/catch` swallows all — never surfaces to user |

**Partial success philosophy:** If 3/5 chunks succeed → their suggestions are applied. If 25/30 suggestions are found → 25 are applied and 5 reported as "not found". Never fail the whole pipeline because of a partial failure.
