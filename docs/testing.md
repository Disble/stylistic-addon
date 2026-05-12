# Testing

This document is the canonical testing reference for `stylistic-addon`.

Complementary documents:

- [`architecture.md`](./architecture.md) — current collaborator boundaries and
  workflow contracts.
- [`linting-and-file-anatomy.md`](./linting-and-file-anatomy.md) — enforced
  `__tests__/` placement and test-support anatomy.
- [`troubleshooting.md`](./troubleshooting.md) — real-host symptoms and current
  operational gotchas.
- [`replace-resolution-postmortem.md`](./replace-resolution-postmortem.md) —
  long-form replace-resolution lessons.

## Core principles

### Real host beats automated GREEN

If Word real and the tests disagree, the tests are wrong or incomplete.

### Test behavior, not internal choreography

Prefer tests that protect:

- final document semantics,
- returned workflow/adapter contracts,
- user-visible taskpane states,
- failure when an invariant is broken.

Weak assertions like “`next()` was called” or “method called once” are only
acceptable when they protect a real contract.

### Use strict mocks at the host boundary

When a bug lives at the Office.js boundary, permissive mocks are worse than no
mock at all. Model:

- proxy loading requirements,
- mutation effects on later queries,
- content-control deletion semantics,
- tracked-change visibility gaps.

## Confidence tiers

| Tier | Meaning | Typical examples |
| --- | --- | --- |
| Tier 1 | Primary regression shield | Word adapters, resolution workflows, strict Office.js contract tests |
| Tier 2 | Useful support coverage | taskpane rendering and user-flow guardrails |
| Tier 3 | Low-signal plumbing | repetitive handler bookkeeping and generic orchestrator flow |

Test count is not confidence. Tier 1 matters most.

## Repository conventions

- All tests under `src/` live in sibling `__tests__/` directories.
- Test helpers/mocks may live in explicit helper modules such as
  `*TestHelper.ts` or `*Mocks.ts`.
- Use Vitest via Bun:
  - `bun run test`
  - `bun run test -- src/taskpane/__tests__/TaskpaneSuggestionResolution.test.ts`
  - `bun run typecheck`

## Bug-fix workflow

When a regression escapes:

1. Capture the real-host symptom first.
2. Identify the semantic mismatch (`actual` vs `expected`).
3. Add or tighten the right regression test.
4. Fix the implementation.
5. Update the relevant docs if the lesson is repo-specific.

### Mock mutation gotcha

When a strict Word mock models `accept()` / `reject()` by removing tracked changes
from the same backing collection, any batch helper that loops the live mutable
array can silently skip the second revision in a replace pair. Iterate over a
snapshot (`[...items]`) when the mutation callback shrinks the collection during
resolution.

## Focused suites

Current taskpane guardrail suites live under:

- `src/taskpane/__tests__/TaskpaneEntrypoint.test.ts`
- `src/taskpane/__tests__/TaskpaneSuggestionPresentation.test.ts`
- `src/taskpane/__tests__/TaskpaneSuggestionResolution.test.ts`
- `src/taskpane/__tests__/TaskpaneFeedback.test.ts`

These are valid tests, but they are still Tier 2 compared with strict adapter
and workflow coverage.
