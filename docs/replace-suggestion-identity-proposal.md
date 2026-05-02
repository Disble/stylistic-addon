# Proposal — Replace Suggestion Identity in Word

## Status

- **State**: Proposed architectural direction
- **Motivation**: recurrent false `already-resolved` regressions in real Word

---

## Problem

The add-in currently treats a `replace` suggestion as if a single Word artifact
could reliably represent the whole review unit.

That assumption is false.

In native Word Track Changes, a logical replace suggestion is semantically
composed of at least two parts:

1. the **added/inserted side**
2. the **deleted/original side**

The current implementation creates a `ContentControl` on the inserted side and
later tries to reconstruct the deleted side through `cc.getTrackedChanges()`,
`body.getTrackedChanges()`, and spatial relations such as overlap or adjacency.

This means the system is not resolving a strong identity. It is inferring one.

---

## Root Cause

The recurrent regression is not primarily a UI problem, a state-machine problem,
or a mediator problem.

The root cause is a **modeling error**:

> the system confuses a partial operational handle in Word with the full domain
> identity of a replace suggestion.

Today, the system can reach this invalid conclusion:

- no tracked changes were observed around the suggestion Content Control
- therefore the suggestion is `already-resolved`

That conclusion is epistemologically unsound.

**Absence of observed tracked changes is not proof of resolution.**

It only proves that the current observation strategy could not locate the full
replace artifact set.

---

## Current Model

### Materialization

- `ApplySuggestionCommand` performs `insertText(..., replace)`
- the add-in then creates a `ContentControl` over `insertedRange`
- `cc.tag` stores the suggestion id/form
- `cc.title` stores the original anchor text

### Resolution

- `WordAdapter` finds the CC by tag
- it queries `cc.getTrackedChanges()`
- it supplements with `body.getTrackedChanges()` filtered by
  `compareLocationWith(ccRange)`
- if zero related tracked changes are found, it returns `already-resolved`

### Why this is structurally weak

Because the CC is attached to the inserted side, while the resolution semantics
depend on both inserted and deleted sides.

The deleted side is therefore not identified directly. It is inferred by host
visibility and geometry.

---

## Why It Regresses

This bug family regresses because improvements in surrounding architecture do not
change the identity model at the Word boundary.

Refactors to the taskpane, workflow, mediator, or state machines can improve
coordination and presentation, but they do not repair this foundational issue:

- the system still treats a partial Word artifact as if it were the suggestion
- the system still promotes failed observation into terminal business truth

As long as that remains true, regressions will keep resurfacing under different
host shapes.

---

## Architectural Principle

For `replace` suggestions, the corrected principle is:

> One replace suggestion has one domain identity and may have multiple Word
> artifact references.

Word artifacts are **references**, not competing identities.

---

## Recommended Direction

## Compound identity model for replace suggestions

Treat a replace suggestion as a composed review unit with:

- one **domain identity**
- one or more **Word artifact references**

At minimum, the model should distinguish:

- the inserted-side reference
- the deleted-side reference
- any operational anchor used to re-find the suggestion in Word

### Desired invariants

1. A replace suggestion is **not fully observable** unless both semantic sides
   can be accounted for.
2. `already-resolved` must require positive confirmation, not absence of
   evidence.
3. A single Content Control on the inserted side must never be treated as the
   whole identity of the replace suggestion.

---

## Immediate Defensive Move

Until the stronger identity model exists, the system should stop interpreting:

- `0 tracked changes observed`

as:

- `already-resolved`

For track-change/replace suggestions, that result should instead be modeled as a
non-terminal, non-confirmatory state such as:

- `unobservable`
- `identity-lost`
- or a conservative retryable error state

This avoids lying to the taskpane and avoids sending feedback for an
unconfirmed resolution.

---

## Domain Impact

If this direction is adopted, the domain should evolve toward explicit
distinctions such as:

- `ReviewSuggestion` as a domain entity
- `WordArtifactRef` as an operational reference type
- `ReplaceSuggestionIdentity` or equivalent composed identity model
- `ObservationStatus` distinct from business resolution status

This implies that current result/status types should eventually stop collapsing:

- confirmed resolution
- unobservability
- missing anchor
- host error

into the same narrow set of terminal outcomes.

---

## Architecture Impact

### `ApplySuggestionCommand`

Should stop being interpreted as the place where replace identity is fully
defined by a single inserted-side Content Control.

### `WordAdapter`

Should separate these concerns more explicitly:

1. locate Word refs
2. observe semantic replace state
3. resolve tracked changes
4. conclude business status only when observation is strong enough

### Workflow / Taskpane

Workflow and taskpane should treat ambiguous observation as ambiguous — not as a
terminal success state.

---

## Star Map (yes, the stars)

```text
                     ✦ deleted-side ref
                    /
                   /
      ✦ domain identity ───────── ✦ inserted-side ref / CC
                   \
                    \
                     ✦ operational anchor / fallback locator

Current bug: the system mistakes the right-hand star for the whole constellation.
Correct model: the constellation is the suggestion; each star is only one ref.
```

---

## Final Rule

**Never upgrade observability failure into terminal resolution.**

That is the core rule this proposal records for future implementation.

---

## Rollout note — hard cut to compound-v2 for replace resolution

The adopted direction no longer preserves `legacy-v1` as a supported resolution
path for replace suggestions.

- new replace suggestions persist `compound-v2` metadata in the Content Control
  title payload,
- replace resolution requires compound-v2 metadata instead of bare-ID fallback,
- if v2 metadata exists but is incomplete/corrupt, the result must be
  `identity-lost` and feedback must be skipped,
- artifacts that cannot satisfy compound-v2 resolution requirements are not
  treated as actionable replace identities.

---

## Production extension — operational-wrapper subtypes

The replace identity correction generalized into a broader operational-wrapper
model for native Track Changes suggestions.

Replace remains the composed two-sided case: inserted/current side plus
deleted/original side plus operational anchor. Delete-only and formatting are not
forced into that shape because Word exposes them differently:

- **delete-only** has no meaningful inserted-side identity; the backend sends
  `suggestedText: ""`, Word performs `insertText("", replace)`, and the wrapper
  range/delete side is the strong evidence scope.
- **formatting** has unchanged reviewed text; the backend sends exact markdown
  `*anchor*` / `**anchor**`, and the Word adapter applies native italic/bold font
  mutations that Word exposes as `Formatted` tracked changes.

The invariant stays the same: the add-in must not treat one incidental Word
artifact as the whole suggestion. Each subtype persists the minimum explicit refs
needed to locate and observe its own native Word evidence.
