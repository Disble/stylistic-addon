---
name: stylistic-addon-testing
description: >
  Testing conventions and anti-patterns for the stylistic-addon Word Add-in project.
  Trigger: When writing, reviewing, or refactoring tests in `stylistic-addon`, especially Office.js mocks, Word adapter behavior, taskpane flows, and regression coverage.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "1.2"
---

# Skill: stylistic-addon-testing

## When to Use

- Writing or changing any `.test.ts` file in `stylistic-addon`
- Touching Office.js mocks, `Word.run`, tracked changes, comments, or content controls
- Reviewing whether a test provides real confidence or only plumbing coverage

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

#### Verified repo lesson: `Range.getTrackedChanges()` can expose replace evidence that CC/body queries miss

Another real-host correction showed that a replace suggestion can remain fully
pending in Word while:

- `ContentControl.getTrackedChanges()` returns no actionable changes,
- `Body.getTrackedChanges()` filtered by proximity also returns nothing useful,
- but `ContentControl.getRange().getTrackedChanges()` exposes the actionable
  insertion/deletion pair.

That means:

- `cc.getTrackedChanges()` is also a **partial source**, not a privileged one,
- body-level proximity heuristics are still insufficient on their own,
- strict tests must model the case where the host exposes the relevant changes
  only through the CC range.

**Testing rule**: add a RED case where `cc.getTrackedChanges()` is empty,
`body.getTrackedChanges()` is empty or irrelevant, and only
`cc.getRange().getTrackedChanges()` returns the actionable replace changes. The
adapter must resolve successfully in that scenario.

#### Verified repo lesson: Deleted-side re-observation must prefer deleted-side evidence over stale CC-range proxies

Another escaped regression showed that after `acceptSuggestion` resolves the
`Added` side of a replace, Word can still expose more than one `Deleted`
candidate for the remaining semantic step:

- a stale `cc.getRange().getTrackedChanges()` proxy that throws or silently no-ops,
- and a fresh `deletedSideRef`-anchored range that still resolves the original text.

If the re-observation path prefers the stale CC-range proxy, the workflow can
apply only the inserted side, leave the original text unresolved, and then fail
closed during post-execute verification.

**Testing rule**: add a RED case where the remaining `Deleted` side is exposed
through both `ccRange` and `deletedSide`, but only the `deletedSide` proxy is
actionable. The adapter must choose `deletedSide` first and complete the
resolution without invoking the stale CC-range proxy.

#### Verified repo lesson: semantic-step recovery must not fall back to full-pair re-observation

Another escaped regression showed that fixing the between-step re-observation is
not sufficient when the second semantic step itself fails.

After `acceptSuggestion` completes the `Added` side, the `Deleted` step can
still fail with a stale proxy. If that recovery path re-observes the full
replace pair instead of re-observing only the remaining semantic side, it can
reintroduce the wrong pair selection and keep the workflow stuck in
`Deleted,Added` even though only `Deleted` should still be under consideration.

**Testing rule**: add a focused regression for `ResolveSuggestionCommand`
showing that failed replace-step recovery calls the side-specific re-observation
helper for the same semantic side and does not route through the generic
full-pair observer.

#### Verified repo lesson: the deleted-side locator must not contribute Added evidence

Another real-host log exposed a more specific observer bug: the range relocated
through `deletedSideRef` can still expose tracked changes, but if that loader is
treated as a generic range source it may return an `Added` and contaminate the
post-execute replace pair as `Deleted(ccRange) + Added(deletedSide)`.

That is semantically invalid. `deletedSideRef` exists to re-anchor the original
text side of the replace, so its contribution must be restricted to `Deleted`
tracked changes.

**Testing rule**: add a RED case where the deleted-side range exposes both
`Added` and `Deleted`, and assert that the observer keeps only the `Deleted`
entry from that source.

#### Verified repo lesson: persisted identity must be used to re-localize actionable ranges

Another escaped regression showed that persisting richer metadata (`compound-v2`)
is not enough if the adapter only validates it and then falls back to CC/body
proximity heuristics.

In real-host-style scenarios, the actionable replace changes may be discoverable
only after re-locating the persisted operational anchor and querying tracked
changes from that range.

That means:

- identity metadata is not merely diagnostic,
- `anchorRef` must be tested as an operational recovery mechanism,
- a suite that validates metadata presence but never exercises anchor-based
  re-location is still incomplete.

**Testing rule**: add a RED case where CC-scoped tracked changes, CC-range
tracked changes, and body-proximity tracked changes are all empty, but the range
re-located through the persisted operational anchor exposes the actionable
replace changes. The adapter must resolve successfully in that scenario.

#### Verified repo lesson: do not delete colocated comments before resolution is confirmed

Another correction showed that the colocated Stylistic comment can still carry
useful host-level evidence for a pending replace suggestion.

If the adapter deletes that comment before it finishes observing the replace,
it can destroy the last actionable range that still exposes the tracked changes.

That means:

- comment cleanup is not always a harmless pre-step,
- `acceptSuggestion` / `rejectSuggestion` sequencing matters,
- tests must model the case where the comment range is the only remaining
  observable source of the replace change pair.

**Testing rule**: add a RED case where CC-scoped, CC-range, operational-anchor,
and body-proximity evidence are empty, but the colocated comment range exposes
the actionable tracked changes. Resolution must succeed, and the comment must be
deleted only after tracked-change resolution is confirmed.

#### Verified repo lesson: duplicate tags can make `getByTag()` return the wrong CC

Another real-host correction showed that `contentControls.getByTag()` may return
multiple artifacts for the same logical suggestion, and blindly taking
`items[0]` can select a stale or legacy wrapper that does not carry the expected
`compound-v2` metadata.

That means:

- tag equality alone is not sufficient to choose the operational artifact,
- tests that assume one tag maps to exactly one CC can hide real-host failures,
- resolution must explicitly prefer the CC whose metadata matches the expected
  replace identity contract.

**Testing rule**: add a RED case where `getByTag()` returns multiple CCs with
the same tag, the first one is legacy/non-v2, and a later one has valid
`compound-v2` metadata plus actionable tracked changes. The adapter must choose
the valid v2 candidate instead of `items[0]`.

#### Verified repo lesson: re-observation must remap logical candidates to fresh proxies

Another escaped regression showed that preserving the previously selected
Content Control by object reference is not enough once Word re-materializes the
 suggestion after a semantic step.

That means:

- re-observation must never inject the old `preferredCc` proxy back into the
  observer,
- the workflow must resolve that preference to a fresh candidate from the new
  `getByTag()` result,
- matching by logical identity (`tag` + persisted `title` metadata) is safer
  than proxy identity.

**Testing rule**: add a RED case where the first observation uses one CC proxy,
the next `getByTag()` returns a fresh equivalent candidate, and the old proxy
throws if reused. The workflow must still resolve successfully through the
fresh candidate.

#### Verified repo lesson: `compound-v2` fixtures must match the real suggestion identity

Another escaped regression showed that tests can still lie even when they use a
`compound-v2` title, if that metadata does not actually match the suggestion
under test.

The concrete failure mode was:

- the test suggestion used one `anchor` / `context`,
- the helper-generated `compound-v2` title kept a default `deletedSideRef.value`
  and `anchorRef.value` from another scenario,
- stricter identity validation then returned `identity-lost`,
- and the suite revealed it had been certifying inconsistent fixtures rather
  than real replace identity.

That means:

- `compound-v2` presence alone is not enough for a trustworthy test,
- fixture metadata must be semantically aligned with `suggestion.id`, tag,
  `suggestion.anchor`, and `suggestion.context`,
- helper defaults are dangerous once the scenario stops using the default
  anchor/context pair.

**Testing rule**: whenever a test uses a non-default replace suggestion, pass an
explicit `ccTitle` / `makeCompoundV2Title(...)` whose `insertedTag`,
`deletedValue`, and `anchorValue` match the exact suggestion fixture. Otherwise
the test is permissive noise, not confidence.

#### Verified repo lesson: reject can succeed in Word before later cleanup/state reads fail

Another real-host correction showed that `rejectSuggestion` can complete the
host-level rejection successfully and still throw a later `GeneralException`
during cleanup or post-resolution state inspection.

If the adapter treats that late exception as total failure, the taskpane becomes
desynchronized from the document: Word is already correct, but the UI reports an
error.

That means:

- not all post-resolution exceptions are semantically equal,
- reject flows must distinguish "reject failed" from "reject succeeded but
  later housekeeping failed",
- tests must explicitly model host success followed by late invalidation.

**Testing rule**: add a RED case where `reject()` is executed on the tracked
changes successfully, but a later cleanup or `inspectDocumentReviewState()` call
throws `GeneralException`. The adapter must still return `rejected` rather than
degrading to `error`.

#### Verified repo lesson: unknown mutation verification is not zero tracked changes

Another real-host correction showed that `accept()` can flush cleanly on a
stale replace-side `TrackedChange` proxy while Word still exposes the same
`Deleted,Added` pair as pending. In that run, the executor's body
tracked-change count probe failed with Word's misnamed `InvalidRibbonDefinition`
error, and the previous implementation returned `0` from the probe failure.

That test strategy failed because it encoded two false assumptions:

- a failed body count probe is equivalent to a document with zero tracked
  changes,
- accept flows are exempt from the stale-proxy silent no-op pattern that had
  previously been seen on reject.

Both assumptions are wrong. Unknown host evidence must remain **unknown**, and
both `accept` and `reject` must treat a non-decreasing known body count as a
silent no-op signal.

**Testing rule**: executor tests must include both (1) accept-side silent no-op
detection when body tracked-change count stays flat after sync, and (2) probe
failure cases where `body.getTrackedChanges()` throws `InvalidRibbonDefinition`.
The report must expose an unverified mutation instead of pretending the count is
`0`, and replace workflow tests must force a fresh re-observation before any
terminal success.

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

### 7. Workflow extraction deserves its own tests

When logic moves out of `taskpane.ts` or `WordAdapter.ts` into a dedicated workflow object (for example `SuggestionResolutionWorkflow`), do not rely only on adapter tests plus taskpane tests.

Add a focused suite for the workflow itself when it owns:

- action branching (`accept` / `reject`)
- non-blocking feedback dispatch
- workflow result semantics such as `feedbackStatus`

Otherwise the workflow can become a blind spot hidden between UI guardrails and adapter contracts.

### 8. Track Changes lifecycle needs both negative and positive proofs

For this repo, it is not enough to prove that comment-only batches do **not** enable Track Changes.

You must also prove the positive path:

- document starts with tracking `off`
- first real `track-change` suggestion is applied
- workflow enables Track Changes exactly once
- result semantics expose that activation (`trackChangesActivatedForBatch`)

Testing only the negative path leaves the new lifecycle partially uncertified.

---

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
