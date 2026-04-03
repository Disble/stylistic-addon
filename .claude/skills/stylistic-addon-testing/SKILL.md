---
name: stylistic-addon-testing
description: >
  Testing conventions and anti-patterns for the stylistic-addon Word Add-in project.
  Trigger: When writing, reviewing, refactoring, or debugging tests in `stylistic-addon`, especially Office.js mocks, Word adapter behavior, taskpane flows, and regression coverage.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "1.1"
---

# Skill: stylistic-addon-testing

## When to Use

- Writing or changing any `.test.ts` file in `stylistic-addon`
- Debugging a regression that escaped despite existing tests
- Touching Office.js mocks, `Word.run`, tracked changes, comments, or content controls
- Reviewing whether a test provides real confidence or only plumbing coverage

## Bug Intake Pattern

This repo's bug reports often arrive as:

- screenshots from real Word or the taskpane
- explicit **current behavior**
- explicit **expected behavior**

Treat that input as first-class evidence.

**Workflow**:
1. extract what the screenshot proves visually
2. write down the semantic mismatch (`actual` vs `expected`)
3. only then inspect code and tests

Do NOT rush into implementation before you understand what the host/UI evidence is actually telling you.

## Critical Patterns

### 0. Real host beats automated GREEN

If Word real and the tests disagree, the tests are wrong or incomplete.

**Rule**:
- GREEN is provisional
- real Word is authoritative
- if Word disproves the model, fix the test model first, then the implementation

This is not optional in `stylistic-addon`.

### 1. Test behavior, not internal choreography

Prefer tests that prove a user-visible or adapter-contract outcome.

Good:
- final document semantics
- returned domain result
- contract with Office.js / SDK boundary
- failure when a required invariant is broken

Weak on its own:
- `next()` called
- constructor called
- method called once
- field copied into `ctx`

Those interaction assertions are acceptable only when they protect a real contract.

---

### 2. The most dangerous tests are permissive mocks

The addon's worst regressions have escaped because mocks modeled a friendlier world than the real host.

Examples of false confidence already observed:

- `document.changeTrackingMode` readable without `load()` + `sync()`
- `cc.getTrackedChanges()` assumed to contain both sides of a replace tracked change
- mocked follow-up searches returning success regardless of whether prior mutations preserved document content

**Rule**: if the production bug lives at the API boundary, the mock must become stricter, not more convenient.

---

### 3. Office.js mocks must encode host semantics

For high-risk tests, model these behaviors explicitly:

| Boundary | Required mock behavior |
|---|---|
| Proxy properties | Throw or fail when read before `load()` + `sync()` |
| Mutations | Affect later queries (`search`, tracked changes, comments, CCs) |
| ContentControl deletion | `delete(true)` keeps content, `delete(false)` removes it |
| Tracked changes | Do NOT assume replace operations expose both insertion and deletion unless the mock proves it |

If a mock cannot model the behavior credibly, the test is lower confidence than it looks.

---

### 4. Large test files are a maintainability smell, not the root cause

Split giant suites because they hide gaps and mix responsibilities.

Recommended split for `WordAdapter`:

- `WordAdapterReadText.test.ts`
- `WordAdapterApplySuggestions.test.ts`
- `WordAdapterAcceptSuggestion.test.ts`
- `WordAdapterRejectSuggestion.test.ts`
- `WordAdapterCleanup.test.ts`
- `WordAdapterNavigateToText.test.ts`

Recommended split for `taskpane`:

- `TaskpaneEntrypoint.test.ts`
- `TaskpaneSuggestionPresentation.test.ts`
- `TaskpaneSuggestionResolution.test.ts`
- `TaskpaneFeedback.test.ts`

But DO NOT confuse file splitting with quality improvement. A small lying test is still a lying test.

### Important correction: taskpane tests DO exist in this repo

Earlier guidance in this project family sometimes treated `src/taskpane/**` as effectively excluded from unit-style coverage because it is DOM-dependent.

That is too absolute for `stylistic-addon`.

Current repo reality:

- `TaskpaneEntrypoint.test.ts`
- `TaskpaneSuggestionPresentation.test.ts`
- `TaskpaneSuggestionResolution.test.ts`
- `TaskpaneFeedback.test.ts`

all provide useful fake-DOM guardrail coverage.

**Correct rule**:
- taskpane tests are valid when they protect presentation semantics and user-flow guardrails,
- but they are **Tier 2**, not the primary integration shield,
- fake-DOM taskpane tests must not be mistaken for real Office/Word boundary confidence.

---

### 5. Protect semantic outcomes for tracked changes

Critical regression lesson:

When accepting a replace suggestion, the suite must protect the semantic outcome:

- the inserted text remains
- the deleted/original text is fully resolved
- the taskpane must not report success if the old text remains unresolved

Do NOT certify this with a mock that simply injects both `Deleted` and `Added` tracked changes into `cc.getTrackedChanges()` unless that is the exact contract being verified.

For `acceptSuggestion` / `rejectSuggestion`, prioritize tests that ask:

> What is the final resolved state of the suggestion?

Not only:

> Which mocked methods were called?

#### Verified repo lesson: `cc.getTrackedChanges()` may be incomplete for replace semantics

In this repo, a real regression showed that accepting a suggestion could confirm the inserted text while leaving the deleted/original text unresolved.

The fix pattern was:

- treat `cc.getTrackedChanges()` as a **partial source**, not the only source
- also inspect `body.getTrackedChanges()`
- keep the tracked changes whose ranges overlap the suggestion Content Control
- resolve the union before returning `accepted` / `rejected`

**Testing rule**: add RED tests where the CC-scoped collection only exposes one side of the replace and the body-level collection exposes the missing side. The test must fail before the fix and pass after it.

#### Verified repo lesson: adjacency matters for native replace resolution

Another real Word regression showed the missing deleted/original side of a replace can be `AdjacentBefore` the inserted-side ContentControl, not only overlapping it.

That means:

- overlap-only collection is too narrow for `acceptSuggestion` / `rejectSuggestion`
- tests that only use `OverlapsBefore` / `OverlapsAfter` can still lie
- the resolution contract is broader than comment-cleanup colocation

**Testing rule**: include at least one RED case with `AdjacentBefore` (and, when relevant, `AdjacentAfter`) when verifying native replace accept/reject semantics.

**Methodology rule**: a GREEN suite is provisional until the behavior is rechecked in real Word. If Word disproves the model, update the tests first, then the implementation.

---

### 6. Reclassify tests by confidence tier

Use this mental model during review:

| Tier | Meaning | Examples |
|---|---|---|
| Tier 1 | Primary regression shield | `ApplySuggestionCommand`, strict `WordAdapter` action tests, `MastraAdapter` contract tests |
| Tier 2 | Useful support coverage | `taskpane` rendering and user-flow tests |
| Tier 3 | Plumbing / consolidatable | repetitive handler tests, observer bookkeeping, generic orchestrator flow |

Test count does NOT equal confidence. Tier 1 matters most.

---

### 7. Use the bug investigation ladder

Investigate escaped regressions in this order unless the evidence clearly skips a layer:

1. **Observed symptom** — screenshots, current behavior, expected behavior
2. **UI/taskpane state** — terminal rendering, action buttons, visible user feedback
3. **Adapter result contract** — returned status, semantic success vs cleanup failure
4. **Creation/resolution mechanics** — WordAdapter, ApplySuggestionCommand, comments, CCs, tracked changes
5. **Host semantics** — Office.js range relations, proxy loading, object invalidation, batch timing

This avoids getting trapped in a single file too early.

---

### 8. Choose test types by bug layer

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

### 9. Split tests surgically during active bug work

Do NOT refactor the whole suite during an active bug unless the user asks.

Instead:
- split only the test files on the current bug path
- reduce cognitive load around the live investigation
- keep the wider suite cleanup as a separate effort

This repo already validated that pattern during the native tracked-change bug fixes.

---

### 10. Document corrections of corrections

When a regression reveals that a prior fix or prior test strategy was incomplete:

1. fix the production bug
2. update the relevant skill or instruction file
3. explain why the previous test or assumption failed
4. record the corrected pattern so future agents do not reintroduce the same mistake

This repo MUST preserve operational learning, not just code diffs.

### 10.1 New lesson: workflow extraction deserves its own tests

When logic moves out of `taskpane.ts` or `WordAdapter.ts` into a dedicated workflow object (for example `SuggestionResolutionWorkflow`), do not rely only on adapter tests plus taskpane tests.

Add a focused suite for the workflow itself when it owns:

- action branching (`accept` / `reject`)
- non-blocking feedback dispatch
- workflow result semantics such as `feedbackStatus`

Otherwise the workflow can become a blind spot hidden between UI guardrails and adapter contracts.

### 10.2 New lesson: Track Changes lifecycle needs both negative and positive proofs

For this repo, it is not enough to prove that comment-only batches do **not** enable Track Changes.

You must also prove the positive path:

- document starts with tracking `off`
- first real `track-change` suggestion is applied
- workflow enables Track Changes exactly once
- result semantics expose that activation (`trackChangesActivatedForBatch`)

Testing only the negative path leaves the new lifecycle partially uncertified.

---

### 11. Why bug 2 was faster than bug 1

The second native tracked-change bug was resolved with fewer user iterations because the workflow improved:

- the first bug taught us not to trust the first GREEN
- the active bug-path tests were split into focused suites
- the tracing moved across adapter + UI layers immediately
- the screenshots were used as semantic evidence, not decoration
- the hypothesis was framed in terms of semantic success vs cleanup failure

**Meta-lesson**: the goal is not just to fix the current bug, but to become faster and sharper on the next one.

## Proven Anti-Patterns in This Repo

### ❌ Anti-pattern: certifying a mock's premise

```typescript
const trackedChanges = [deletedTc, addedTc];
cc.getTrackedChanges = () => trackedChanges;

await adapter.acceptSuggestion(suggestion);

expect(deletedTc.accept).toHaveBeenCalled();
expect(addedTc.accept).toHaveBeenCalled();
```

This only proves that if the mock gifts both tracked changes to the adapter, both are accepted.
It does NOT prove the real host exposes both changes that way.

### ✅ Better direction

Use stricter mocks or stateful harnesses that represent what the host actually exposes, then assert the resolved outcome.

---

### ❌ Anti-pattern: too many low-signal plumbing tests

Examples:
- handler tests that only assert `ctx.foo = value`
- orchestrator tests that only assert `next()` chains
- observer tests that exhaustively validate bookkeeping with little product risk

These are not worthless, but they must not dominate the suite budget.

## Commands

```bash
npx vitest run src/adapters/word/ApplySuggestionCommandSearch.test.ts src/adapters/word/ApplySuggestionCommandContentControl.test.ts src/adapters/word/ApplySuggestionCommandTrackingMode.test.ts
npx vitest run src/adapters/word/WordAdapterApplySuggestions.test.ts src/adapters/word/WordAdapterAcceptSuggestion.test.ts src/adapters/word/WordAdapterRejectSuggestion.test.ts src/adapters/word/WordAdapterReadText.test.ts src/adapters/word/WordAdapterAppliedTexts.test.ts src/adapters/word/WordAdapterCleanup.test.ts src/adapters/word/WordAdapterNavigateToText.test.ts
npx vitest run src/taskpane/TaskpaneEntrypoint.test.ts src/taskpane/TaskpaneSuggestionPresentation.test.ts src/taskpane/TaskpaneSuggestionResolution.test.ts src/taskpane/TaskpaneFeedback.test.ts
```

## Resources

- Root audit artifacts:
  - `openspec/changes/addon-test-confidence-audit/exploration.md`
  - `openspec/changes/addon-test-confidence-audit/proposal.md`
- Project instructions:
  - `stylistic-addon/AGENTS.md`
  - `stylistic-addon/CLAUDE.md`
