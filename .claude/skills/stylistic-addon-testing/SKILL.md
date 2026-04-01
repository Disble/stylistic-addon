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

## Critical Patterns

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

- `WordAdapter.getTextToAnalyze.test.ts`
- `WordAdapter.applySuggestions.test.ts`
- `WordAdapter.acceptSuggestion.test.ts`
- `WordAdapter.rejectSuggestion.test.ts`
- `WordAdapter.cleanupResolvedComments.test.ts`
- `WordAdapter.navigateToText.test.ts`

Recommended split for `taskpane`:

- `taskpane.bootstrap.test.ts`
- `taskpane.results-rendering.test.ts`
- `taskpane.accept-reject.test.ts`
- `taskpane.feedback.test.ts`

But DO NOT confuse file splitting with quality improvement. A small lying test is still a lying test.

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

### 7. Document corrections of corrections

When a regression reveals that a prior fix or prior test strategy was incomplete:

1. fix the production bug
2. update the relevant skill or instruction file
3. explain why the previous test or assumption failed
4. record the corrected pattern so future agents do not reintroduce the same mistake

This repo MUST preserve operational learning, not just code diffs.

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
npx vitest run src/adapters/word/ApplySuggestionCommand.test.ts
npx vitest run src/adapters/word/WordAdapter.test.ts
npx vitest run src/taskpane/taskpane.test.ts
```

## Resources

- Root audit artifacts:
  - `openspec/changes/addon-test-confidence-audit/exploration.md`
  - `openspec/changes/addon-test-confidence-audit/proposal.md`
- Project instructions:
  - `stylistic-addon/AGENTS.md`
  - `stylistic-addon/CLAUDE.md`
