# Tasks: Operational Wrapper Replace Resolution

> Historical task artifact. Paths and commands below have been normalized where
> useful so future readers can map them to the current repo layout.

## Phase 1: Audit and cutover plan

- [x] 1.1 Audit `src/adapters/word/ApplySuggestionCommand.ts`, `ReplaceIdentityParser.ts`, `resolution/SuggestionLocator.ts`, `SuggestionResolutionObserver.ts`, `ReplaceResolutionOrchestrator.ts`, `SuggestionResolutionCleanup.ts`, and `cleanup/CommentCleanup.ts`; list every legacy replace path to delete: duplicate scoring/ranking, full-body fallback, semantic-side-only grouping, unsafe replace cleanup, and dormant compatibility branches.
- [x] 1.2 Map every caller/test to rewrite or delete: `ResolveSuggestionCommand.ts`, Word adapter accept/reject/navigation tests, taskpane resolution tests, and parser/locator/observer/cleanup suites; fail the audit if any legacy path has no deletion/replacement owner.
- [x] 1.3 Treat `D:\dev\eln\stylistic\reject-word-testing` as read-only evidence only; extract semantic rules from Scenario 19-23, but do not copy lab UI, buttons, logging scaffolding, or create a production dependency on lab files.

## Phase 2: Strict identity and locate

- [x] 2.1 Extend `src/domain/types.ts` with operational wrapper/group metadata and `ResolutionAbortReason`; add documentation comments to every new/rewritten function, class, and complex logic block touched by this change.
- [x] 2.2 Rework `src/adapters/word/ApplySuggestionCommand.ts` to persist operational wrapper/group identity and abort before mutation on non-unique location; delete replace full-body rescue, tie-break scoring, and replace-specific wrapper reinsertion fallback.
- [x] 2.3 Replace `src/adapters/word/ReplaceIdentityParser.ts` and `src/adapters/word/resolve-suggestion/SuggestionLocator.ts` with strict wrapper validation/exact lookup that returns explicit `ambiguous-location` or `identity-lost`, never ranked fallback.

## Phase 3: Group resolution architecture

- [x] 3.1 Create `src/adapters/word/resolve-suggestion/OperationalWrapperGroupResolver.ts` to expand only explicit contiguous/adjacent Stylistic wrapper groups; do not generalize Scenario 22 into a universal distance heuristic.
- [x] 3.2 Refactor `SuggestionResolutionObserver.ts`, `ResolutionContext.ts`, and `ReplaceResolutionOrchestrator.ts` so operational wrapper lineage is the semantic identity, comments/inserted-side CCs are metadata/UX only, and grouped resolution is all-or-nothing `acceptAll()`/`rejectAll()`.
- [x] 3.3 Make mixed decisions inside one contiguous group degrade explicitly to `unobservable`/`mixed-group` before executor mutation; no silent fallback to per-side, per-fragment, or legacy semantic ranking.
- [x] 3.4 Update `ResolveSuggestionCommand.ts` and `TrackedChangeResolutionExecutor.ts` wiring so ambiguous-location and mixed-group aborts happen before mutation, while unrelated non-Stylistic tracked changes remain untouched.

## Phase 4: Cleanup and legacy deletion

- [x] 4.1 Modify `src/adapters/word/resolve-suggestion/SuggestionResolutionCleanup.ts` and `src/adapters/word/cleanup/CommentCleanup.ts` to enforce comments-only cleanup after replace resolution and preserve wrapper + inserted-side CC metadata.
- [x] 4.2 Delete replace-specific `ContentControl.delete(false/true)` cleanup paths and any dormant helper/branch that can remove replace wrappers/CCs or re-enter the legacy workflow.
- [x] 4.3 Delete or rewrite tests that bless duplicate ranking, body fallback, semantic-side-only replace grouping, or unsafe metadata deletion.

## Phase 5: Spec-driven tests

- [x] 5.1 Add/replace focused adapter tests for Scenario 19 ambiguous abort and Scenario 22 one-word-gap independent wrappers, explicitly documenting that Scenario 22 is narrow evidence and NOT a universal heuristic.
- [x] 5.2 Add/replace grouped-resolution tests for Scenario 20 contiguous grouped `rejectAll()`, Scenario 23 contiguous grouped `acceptAll()`, and explicit mixed-group degradation.
- [x] 5.3 Add/replace cleanup/preservation tests for Scenario 21 comments-only cleanup, metadata persistence, metadata-vs-identity distinction, and preservation of non-Stylistic tracked changes.
- [x] 5.4 Update `src/adapters/word/__tests__/WordAdapterAcceptSuggestion.test.ts`, `src/adapters/word/__tests__/WordAdapterRejectSuggestion.test.ts`, `src/taskpane/__tests__/TaskpaneSuggestionResolution.test.ts`, and related domain tests to assert semantic Word behavior, not permissive implementation trivia.

## Phase 6: No-legacy verification

- [x] 6.1 Run static negative checks proving no reachable `scoreCompoundReplaceIdentityMatch`, `rankResolutionContentControls`, replace full-body fallback, semantic-side-only replace grouping, or replace CC deletion cleanup remains.
- [x] 6.2 Run focused verification only with standard non-build commands: `bun run typecheck` and `bun run test -- src/adapters/word/__tests__/WordAdapterAcceptSuggestion.test.ts src/adapters/word/__tests__/WordAdapterRejectSuggestion.test.ts src/adapters/word/resolve-suggestion/__tests__/SuggestionLocator.test.ts src/adapters/word/resolve-suggestion/__tests__/SuggestionResolutionObserver.test.ts src/adapters/word/resolve-suggestion/__tests__/SuggestionResolutionCleanup.test.ts src/taskpane/__tests__/TaskpaneSuggestionResolution.test.ts`.
- [x] 6.3 Perform call-site review, test review, and diff accountability listing every deleted/replaced legacy file/function/branch; FAIL this change if any dormant fallback, hidden compatibility path, or legacy-behavior test remains.

## Apply correction notes — 2026-04-25

- Verified tasks 3.1 / 3.2 / 5.2 against implementation after adding actual `Range.compareLocationWith()` adjacency validation in `OperationalWrapperGroupResolver` and grouped accept/reject semantic tests.
- Verified task 3.3 after adding a production `mixed-group` path for incompatible per-member evidence inside a contiguous operational wrapper group, with a command-level no-mutation test.
- Verified task 5.4 after adding workflow/taskpane coverage for `ambiguous-location` and `mixed-group` feedback/UI semantics.
- Strengthened task 5.3 cleanup evidence with a shared cleanup-harness test that proves this repo's Office mock only mutates the located Stylistic comment while wrapper CCs, inserted-side CCs, and foreign tracked changes remain present. This is stronger than detached spies, but still mock-model evidence rather than full real-Word proof.
- Focused verification passed with `bun run typecheck` and the non-build Vitest suites listed in the apply summary.
