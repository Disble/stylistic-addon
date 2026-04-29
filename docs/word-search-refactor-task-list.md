# Word Search Refactor Task List

This document defines the approved execution order for the next architecture
cycle. It is intentionally prescriptive. The project already knows what happens
when a cross-cutting capability grows without a clear boundary: the code drifts,
the docs drift, and the bug surface becomes asymmetric. This plan exists to stop
that pattern.

The work is split into two major tracks:

1. **architectural validations first**, because boundaries must be enforced
   before the refactor starts,
2. **new architecture implementation by bounded phases**, because the target
   architecture must arrive through small, testable SDD increments instead of a
   single large rewrite.

---

## 1. Non-negotiable execution rules

### 1.1 Validations ship before feature refactor work

The first implementation step is the new validation layer. That work must be
committed on its own as a legal, complete commit. If the new validations fail
against the existing codebase, that is acceptable and even desirable. A failing
guard proves that the guard is catching the architectural violation it was
designed to detect.

This means the project must resist the temptation to “fix the code first and add
the guard later”. That sequence is backwards. If the guard arrives late, the
refactor has no protection and drift can re-enter immediately.

### 1.2 The architecture refactor is phase-based

The second track is the implementation of the new architecture, but never as a
single big-bang change. Each step must be a bounded SDD with a narrow scope,
clear acceptance criteria, and explicit tests.

### 1.3 TDD is mandatory

Each SDD phase must be developed through TDD. The goal is not just green tests;
the goal is evidence that the new boundary behaves as intended and that the old
bug path is covered by the new design.

### 1.4 Documentation is part of each phase

Each phase must update:

- code-level documentation comments,
- diagrams when boundaries move,
- architecture-aligned wording in the relevant files,
- any tests or helper docs that describe the previous behavior.

No phase is complete if the architecture document and the code comments disagree.

---

## 2. Phase 1 — Architectural validations first

### Objective

Add the new validations before the refactor starts so the future boundary is
enforced mechanically instead of relying on memory or discipline alone.

### Expected commit behavior

This phase is allowed to produce a legal commit whose validations fail on the
current codebase. That outcome is not a defect in the phase. It is the evidence
that the rules detect the exact architecture drift they were introduced to stop.

### Tasks

- [x] 1.1 Define the authorized search boundary in documentation and naming. The
      goal is to lock down the intended ownership of `TextSearchCore`,
      `WordTextLocatorAdapter`, and the future resolution collaborators before
      writing the lint rules. Touches: `docs/architecture.md`, `biome.json`
      planning notes, and the future lint naming strategy.

- [x] 1.2 Add the concrete text-search implementation guard. The goal is to block
      direct imports of concrete search implementations from unauthorized layers
      and force consumers toward ports or authorized composition points. Touches:
      `lints/`, `biome.json`.

- [x] 1.3 Add the canonical search primitives guard. The goal is to block
      redefinitions of canonical helper names outside the approved text-search
      core so the repo cannot silently grow another weakened copy of the search
      engine. Touches: `lints/`, `biome.json`.

- [x] 1.4 Add the taskpane/workflow consumption guard. The goal is to prevent the
      presentation layer from importing concrete search engines or Word-specific
      location helpers directly. Touches: `lints/`, `biome.json`.

- [x] 1.5 Review the complexity ratchet after the new lint rules land. The goal
      is not yet to reduce `ResolveSuggestionCommand.ts`, but to prepare the
      complexity guard for the upcoming split and ensure the exception is treated
      as temporary architectural debt. Touches: `scripts/checkComplexity.mjs`.

- [x] 1.6 Update the architecture documentation to explain the new guardrail
      model and expected failure mode of the first validation commit. Touches:
      `docs/architecture.md`, this task list.

### Phase 1 result snapshot

This phase is now implemented and verified.

- `no-concrete-text-search-import.grit` compiles and is wired through
  `biome.json`.
- `no-search-primitive-redefinition.grit` compiles and now catches the current
  duplicated primitives inside `src/adapters/word/ResolveSuggestionCommand.ts`.
- `no-search-internals-in-taskpane.grit` compiles and is wired through
  `biome.json`.
- `npm run lint` currently fails for the expected reason: the canonical
  primitives guard flags `removeWhitespaceWithIndices` and
  `findWhitespaceInsensitiveSlice` inside `ResolveSuggestionCommand.ts`.
- `scripts/checkComplexity.mjs` was ratcheted down so
  `ResolveSuggestionCommand.ts` remains an explicit temporary exception rather
  than normalized debt.

That failure is intentional architectural evidence, not an accidental breakage.
The repo is now protected against adding more drift before the refactor phases
begin.

### Design justification

This phase uses the architecture-enforcement pattern: define the boundary, then
make the toolchain defend it. Without that sequence, the codebase depends on
tribal memory.

### TDD / verification expectation

This phase is validated by rule behavior, targeted smoke checks, and by proving
that current violations are detected. The important result is not “all green”; it
is “the rule catches the forbidden dependency”.

---

## 3. Phase 2 — SDD-1: Extract the pure text-search core

### Objective

Create the first real architectural boundary by extracting the reusable, pure
search capability out of the current Word-only helper arrangement.

### Tasks

- [x] 2.1 Create the pure text-search module (`domain/text-search/` or
      `core/text-search/`) with explicit types and documentation comments.
      Touches: new core files plus any required domain-level types/ports.

- [x] 2.2 Move canonical normalization and locator helpers into the new core.
      The goal is to establish one authoritative implementation for smart quote
      handling, diacritic handling, field-code skipping, whitespace-insensitive
      matching, locator shortening, and first-alphanumeric fallback derivation.

- [x] 2.3 Add pure unit tests for the core. These tests must prove behavior
      without Office.js mocks.

- [x] 2.4 Update documentation comments and architecture references so the repo
      has one canonical vocabulary for the new capability.

### Phase 2 progress snapshot

The second bounded slice of SDD-1 is now in place, which completes the pure-core
extraction goal for this phase.

- `src/core/text-search/TextSearchCore.ts` now owns all five canonical search
  primitives: `normalizeChar`, `removeWhitespaceWithIndices`,
  `findWhitespaceInsensitiveSlice`, `findUniqueLocatorSubstring`, and
  `findFirstAlphanumericOffset`.
- `src/core/text-search/TextSearchCore.test.ts` now covers both locator
  shortening and first-alphanumeric fallback derivation as pure behavior.
- The transitional `src/adapters/word/WordSearchAdapter.ts` wrapper has since
  been removed as Phase 6 hygiene after all callers moved to `TextSearchCore`
  or `WordTextLocatorAdapter`.
- The Word-facing search path is now expressed directly as `TextSearchCore` +
  `WordTextLocatorAdapter`, without a barrel-like shim in the middle.

### Design justification

This phase applies the separation-of-concerns principle and the ports-and-
adapters philosophy at a finer grain. The capability is cross-cutting and pure;
therefore it deserves its own pure module.

### TDD focus

Write tests around the pure helper behavior first. The extraction is correct only
if the matching rules are proven independently from the host adapter.

---

## 4. Phase 3 — SDD-2: Introduce the Word text locator adapter

### Objective

Bridge the pure search core to Office.js through a dedicated adapter so Word
execution behavior stops leaking across multiple commands.

### Tasks

- [x] 3.1 Define the port or adapter contract for text location requests and
      results. The contract must be explicit enough to support apply, navigation,
      and resolution without leaking UI concerns.

- [x] 3.2 Implement `WordTextLocatorAdapter` using the pure core plus Office.js
      execution semantics.

- [x] 3.3 Migrate `ApplySuggestionCommand` to the new adapter while preserving
      current behavior.

- [x] 3.4 Migrate `WordAdapter.navigateToText()` to the same capability so apply
      and navigation stop owning separate fallback logic.

- [x] 3.5 Add adapter contract tests and migration regression tests.

### Phase 3 progress snapshot

Phase 3 is now complete.

- `WordTextLocatorContext.ts` defines the narrow contracts (`TextLocator`,
  `WordTextLocationRequest`, `WordSearchContainer`) plus the temporary default
  locator composition helper.
- `WordTextLocatorAdapter.ts` exists, is covered by dedicated tests, and owns
  the approved exact / relaxed / pure-core fallback strategy.
- `ApplySuggestionCommand.ts` no longer owns its own fallback search pipeline;
  it now depends on `TextLocator` by injection.
- `BatchApplyOrchestrator.ts` is the current apply composition point for the
  default locator.
- Navigation no longer calls the injected `TextLocator` as a loose fallback by
  itself. It routes suggestion fallback through `SuggestionTextRangeLocator`,
  which composes the same `TextLocator` but enforces `context -> anchor` scope
  and forbids global anchor search.
- The navigation regression suite now proves both the default path and the
  injected-locator path, including the no-global-anchor regression that prevents
  selecting unrelated TOC/heading occurrences.
- The next bounded slice should move into Phase 4 and start splitting
  `ResolveSuggestionCommand` from the outside in.

### Design justification

This phase uses Adapter and Strategy/Policy composition. The pure core decides
how to normalize and derive fallback candidates; the Word adapter decides how to
execute those candidates against Office.js.

### TDD focus

Start with the behavior that must stay stable: current robust apply and
navigation matching. The adapter is correct only if it preserves those strengths.

---

## 5. Phase 4 — SDD-3: Split `ResolveSuggestionCommand`

### Objective

Turn `ResolveSuggestionCommand` into a thin orchestrator and move the real
responsibilities into collaborators.

### Tasks

- [x] 4.1 Extract `ReplaceIdentityParser` so identity parsing and validation stop
      living inside the command file.

- [x] 4.2 Extract `DocumentReviewStateInspector` so before/after state snapshots
      become an explicit collaborator.

- [x] 4.3 Extract `SuggestionLocator` so content-control lookup, ranking, and
      colocated comment discovery become a dedicated responsibility.

- [x] 4.4 Extract `SuggestionResolutionObserver` so tracked-change evidence
      gathering and ambiguity classification become a dedicated responsibility.

- [x] 4.5 Extract `SuggestionResolutionCleanup` and
      `TrackedChangeResolutionExecutor` so mutation and cleanup policy are no
      longer mixed with observation semantics.

- [x] 4.6 Extract `ResolveSuggestionResultFactory` and simplify
      `ResolveSuggestionCommand` to orchestration only.

- [x] 4.7 Route resolution search through the shared text-location capability so
      resolve stops using a weaker fallback than apply/navigation.

- [x] 4.8 Add focused tests for each collaborator and regression tests for the
      escaped bug paths.

### Phase 4 is now complete

- `ReplaceIdentityParser.ts` owns replace-identity parsing/validation.
- `DocumentReviewStateInspector.ts` owns document snapshot reading and
  reject-tolerant post-resolution inspection.
- `SuggestionLocator.ts` owns CC lookup, ranking, and colocated comment
  discovery.
- `SuggestionResolutionObserver.ts` owns evidence gathering and now routes
  operational-anchor search through the shared text-location capability.
- `SuggestionResolutionCleanup.ts` and `TrackedChangeResolutionExecutor.ts`
  separate cleanup policy from mutation policy.
- `ResolveSuggestionResultFactory.ts` owns stable taskpane-facing result
  shaping.
- `CommentOnlySuggestionResolver.ts` owns the comment-only branch.
- `ResolveSuggestionCommand.ts` is now a thin orchestrator instead of a giant
  mixed-responsibility command.

Validation evidence for the completed phase:

- `npx vitest run "src/adapters/word/WordAdapterAcceptSuggestion.test.ts" "src/adapters/word/WordAdapterRejectSuggestion.test.ts"` → green (`40 tests`).
- Focused collaborator suites are green:
  - `ReplaceIdentityParser.test.ts`
  - `SuggestionLocator.test.ts`
  - `DocumentReviewStateInspector.test.ts`
  - `ResolveSuggestionResultFactory.test.ts`
- Focused lint for the new resolution collaborators is green after converting
  collaborator-only imports to `import type`.
- `npm run check:filenames` is green with the expanded collaborator suffix
  vocabulary.

Important architectural result:

- The canonical text-search primitive guard is no longer red because of
  `ResolveSuggestionCommand.ts`. Resolution now consumes the shared search
  capability through `SuggestionResolutionObserver` + `TextLocator`, matching
  the apply/navigation path.

### Design justification

This phase applies Orchestrator + Collaborators, Command, Result Factory, and
single-responsibility decomposition. The main gain is epistemic clarity: locate,
observe, mutate, clean up, and report are different responsibilities and must be
treated as such.

### TDD focus

The tests must prove that the split does not reintroduce false terminal states,
does not degrade `compound-v2` handling, and keeps reject-side cleanup behavior
conservative when Word invalidates proxies.

---

## 6. Phase 5 — SDD-4: Align contracts, UI semantics, and stale states

### Objective

Clean up the remaining semantic drift after the structural refactor, especially
around status modeling and UI expectations.

### Tasks

- [x] 5.1 Remove `already-resolved` from the active contract. The runtime,
      workflow, state machine, taskpane, and tests now treat it as a removed
      branch rather than a produced or reserved state.

- [x] 5.2 Update workflow, mediator, and taskpane semantics to match the final
      resolution contract after the split.

- [x] 5.3 Tighten tests for feedback suppression, retryable states, and explicit
      warning states.

- [x] 5.4 Update documentation comments, architecture diagrams, and any related
      docs to reflect the new final state.

### Phase 5 is now complete

The active contract no longer contains `already-resolved` anywhere in runtime
code, UI semantics, workflow branching, or tests.

- `SuggestionActionResult.status` now represents only:
  `accepted`, `rejected`, `unobservable`, `identity-lost`, `cc-not-found`,
  `not-found`, and `error`.
- `SuggestionResolutionWorkflow` emits feedback only for `accepted` and
  `rejected`.
- `SuggestionStateMachine`, taskpane rendering, and progress summary semantics
  no longer reserve a hidden branch for `already-resolved`.
- Ambiguous host knowledge continues to degrade conservatively to
  `unobservable` or `identity-lost`.

Validation focus for this phase:

- contract/type tests prove the branch is gone,
- taskpane resolution tests still protect warning states and feedback
  suppression,
- documentation now describes the final contract instead of the transitional
  one.

### Design justification

This phase is a contract-alignment phase. Structural refactors often leave stale
states behind. That is acceptable only temporarily. The architecture must end in
semantic consistency, not just prettier file boundaries.

### TDD focus

Drive this phase from contract and UI tests. The important question is not “did
we refactor?” but “does the system now tell the truth about what it knows and
what it cannot prove?”

---

## 7. Phase 6 — Continuous documentation and code hygiene

### Objective

Make documentation and implementation stay synchronized across all phases.

### Tasks

- [x] 6.1 Ensure every new class, function, and non-trivial branch has updated
      documentation comments.

- [x] 6.2 Update architecture diagrams whenever a boundary becomes real in code.

- [x] 6.3 Remove obsolete comments and outdated explanations that describe the
      pre-refactor structure.

- [x] 6.4 Review helper names, module names, and folder names so the code speaks
      the same architectural language as the docs.

### Phase 6 is now complete

This hygiene phase removed the remaining barrel-like wrapper that no longer
carried architectural weight and aligned the surrounding docs/comments with the
real post-refactor structure.

- The transitional `src/adapters/word/WordSearchAdapter.ts` shim and its wrapper
  test were removed once all production callers had already migrated away.
- `TextSearchCore` and `WordTextLocatorAdapter` are now the explicit search
  vocabulary with no compatibility barrel in between.
- `WordTextLocatorContext.ts` was reviewed and retained intentionally because it
  still owns real shared contracts plus the authorized default-locator wiring;
  it is transitional, but not an empty barrel.
- Architecture diagrams, file maps, and task-list snapshots now describe the
  actual module graph instead of the intermediate migration state.

Validation focus for this phase:

- pure/core and Word-facing locator tests remain green after removing the shim,
- focused lint remains green after deleting stale Grit references,
- `check:filenames` remains green after the source tree shrink.

### Design justification

Documentation debt is architecture debt. If the file map and the diagrams lie,
the next refactor starts from false assumptions.

---

## 8. Definition of done for the overall initiative

This initiative is done only when all of the following are true:

1. the validation layer exists and enforces the intended boundaries,
2. text search has one canonical pure core,
3. Office.js search execution lives behind a dedicated adapter,
4. `ResolveSuggestionCommand` is an orchestrator instead of a god object,
5. apply, navigation, and resolution all use the same text-location capability,
6. tests prove the bug path and the new boundaries,
7. code comments and `docs/architecture.md` reflect reality.

Anything less may be progress, but it is not the finished architecture.
