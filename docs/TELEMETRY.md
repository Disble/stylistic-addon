# Telemetry

This document defines the observability mindset for `stylistic-addon`.

It exists because real Word regressions have shown that a clean architecture is not enough on its own. The add-in also needs structured evidence about **what phase ran**, **what Word already mutated**, and **what the UI is allowed to claim afterward**.

---

## 1. Core principles

### 1.1 Semantic truth beats housekeeping failure

If Word already resolved a suggestion semantically, later failures in cleanup, telemetry, or post-action inspection MUST NOT downgrade that truth into a retryable error.

Telemetry exists to expose that distinction, not to blur it.

### 1.2 Telemetry is best effort

Telemetry sinks MAY fail. Those failures MUST NOT change workflow semantics.

- Do log or collect a warning locally.
- Do NOT turn a semantic success into `error` because telemetry emission failed.

### 1.3 Structured events beat ad-hoc logs

Free-form `console.log` messages are useful during local debugging, but they are not enough as a long-term contract.

New workflow observability SHOULD prefer a structured event envelope with stable field names.

### 1.4 One workflow attempt, one correlation id

Every apply or resolution attempt SHOULD have a stable `workflowAttemptId` so events from different phases can be correlated later.

At minimum, events SHOULD include:

- `workflowAttemptId`
- `suggestionId`
- `action` (`apply`, `accept`, `reject`, etc.)
- `phase`
- `outcome`

### 1.5 Warnings are not errors

The system must distinguish:

- **terminal success**,
- **terminal success with warnings**,
- **retryable error**.

If a workflow mixes those concepts, the taskpane will lie.

---

## 2. Event envelope

Recommended minimal event shape:

```ts
interface TelemetryEvent {
  workflowAttemptId: string;
  suggestionId?: string;
  action: string;
  phase: string;
  outcome: "started" | "completed" | "failed" | "reconciled";
  metadata?: Record<string, string | number | boolean>;
}
```

Guidelines:

- keep field names stable,
- prefer booleans / enums over prose when possible,
- reserve `metadata` for phase-specific details,
- never hide the phase where the failure happened.

---

## 3. Resolution workflow observability

For suggestion resolution, telemetry SHOULD map to explicit workflow phases:

1. `locate`
2. `observe-before`
3. `execute`
4. `cleanup-comment`
5. `cleanup-metadata`
6. `cleanup-anchor`
7. `inspect-after`

Why this matters:

- `locate` tells us whether the right Word artifact was found and whether
  duplicate/corrupt identity metadata affected selection,
- `execute` tells us what the host mutation attempted,
- `cleanup-comment` tells us whether the colocated Stylistic comment was removed
  or whether Word invalidated the comment proxy after a successful mutation,
- `cleanup-metadata` tells us whether resolved inserted-side and operational
  wrapper Content Controls were removed before the final state snapshot,
- `cleanup-anchor` tells us whether comment-only anchors were removed while
  preserving visible text,
- `inspect-after` tells us whether the final document-derived state was observable.

If these phases collapse into one generic `error`, debugging becomes guesswork.

`cleanup-metadata` metadata must stay primitive because the current event
contract is `Record<string, string | number | boolean>`. Prefer counts and
serialized tag lists such as `deletedContentControlCount` and
`deletedContentControls` over raw arrays.

---

## 4. Apply/search observability

Text search bugs in this add-in have shown that Word may:

- fail to find a stale backend context,
- find only a partial context fragment,
- return multiple anchor candidates across heading/body paragraphs.

For search-heavy workflows, telemetry SHOULD capture:

- whether the match came from exact, relaxed, or fallback search,
- whether search escalated from context to paragraph or full body,
- whether multiple candidates existed,
- how the final candidate was disambiguated.

This is especially important for “No encontrado” and mislocation regressions.

---

## 5. Non-interference rules

Telemetry MUST follow these rules:

1. MUST NOT throw into product workflows.
2. MUST NOT decide business status.
3. MUST NOT overwrite semantic reconciliation.
4. SHOULD surface warning context when cleanup or inspection fails.
5. MAY use a console-backed adapter first, but the event contract must remain reusable for future sinks.

---

## 6. Testing guidance

Telemetry work still needs real guardrails.

When adding or refactoring observability:

- write tests for semantic outcomes first,
- then assert the structured telemetry facts that matter,
- avoid permissive mocks that “gift” successful telemetry paths,
- add regressions for terminal-with-warning behavior, not only happy-path success.

If Word real disproves the test, update the model before trusting the GREEN.

---

## 7. Checklist for future features/refactors

Before merging a workflow change, ask:

- What is the semantic source of truth?
- Which phases can fail after semantic success?
- What warning should survive that failure?
- Which correlation id ties this attempt together?
- What structured event would let us debug this later without screenshots?

If those answers are missing, the workflow is under-observed.
