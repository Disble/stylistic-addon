---
name: stylistic-addon-debugging
description: >
  Bug investigation workflow and fix methodology for the stylistic-addon Word Add-in project.
  Trigger: When debugging a regression that escaped tests, investigating Word/Office.js behavior mismatches, or fixing a bug in `stylistic-addon`.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "1.0"
---

# Skill: stylistic-addon-debugging

## When to Use

- Debugging a regression that escaped despite existing tests
- Word real-host behavior contradicts what the test suite suggests
- Starting a bug investigation in this repo

## Bug Intake Pattern

This repo's bug reports often arrive as:

- screenshots from real Word or the taskpane
- explicit **current behavior**
- explicit **expected behavior**

Treat that input as first-class evidence.

**Workflow**:
1. Extract what the screenshot proves visually
2. Write down the semantic mismatch (`actual` vs `expected`)
3. Only then inspect code and tests

Do NOT rush into implementation before you understand what the host/UI evidence is actually telling you.

## Investigation Workflow

### 1. Use the bug investigation ladder

Investigate escaped regressions in this order unless the evidence clearly skips a layer:

1. **Observed symptom** — screenshots, current behavior, expected behavior
2. **UI/taskpane state** — terminal rendering, action buttons, visible user feedback
3. **Adapter result contract** — returned status, semantic success vs cleanup failure
4. **Creation/resolution mechanics** — WordAdapter, ApplySuggestionCommand, comments, CCs, tracked changes
5. **Host semantics** — Office.js range relations, proxy loading, object invalidation, batch timing

This avoids getting trapped in a single file too early.

---

### 2. Choose test types by bug layer

For cross-layer bugs, do NOT spray the same assertion everywhere. Use the right test in the right place:

| Layer | Test type | Purpose |
|---|---|---|
| Creation / host interaction | causal or contract test | prove the real boundary behavior |
| Adapter | regression + semantic verification | prove the bug is fixed semantically |
| Taskpane / UI | guardrail test | prevent false-success or retryable bad states |

Useful categories for this repo:
- **regression tests** — reproduce the escaped bug exactly
- **semantic/state verification tests** — verify resolved document or UI state
- **contract/boundary tests** — Office.js/SDK expectations
- **guardrail tests** — stop misleading terminal UI states
- **bug-path focused split** — isolate only the suites involved in the active bug

---

### 3. Split tests surgically during active bug work

Do NOT refactor the whole suite during an active bug unless the user asks.

Instead:
- split only the test files on the current bug path
- reduce cognitive load around the live investigation
- keep the wider suite cleanup as a separate effort

This repo already validated that pattern during the native tracked-change bug fixes.

---

### 4. Document corrections of corrections

When a regression reveals that a prior fix or prior test strategy was incomplete:

1. Fix the production bug
2. Update the relevant skill or instruction file
3. Explain why the previous test or assumption failed
4. Record the corrected pattern so future agents do not reintroduce the same mistake

This repo MUST preserve operational learning, not just code diffs.

---

### 5. Meta-lesson: why bug 2 was faster than bug 1

The second native tracked-change bug was resolved with fewer user iterations because the workflow improved:

- the first bug taught us not to trust the first GREEN
- the active bug-path tests were split into focused suites
- the tracing moved across adapter + UI layers immediately
- the screenshots were used as semantic evidence, not decoration
- the hypothesis was framed in terms of semantic success vs cleanup failure

**Goal**: not just to fix the current bug, but to become faster and sharper on the next one.
