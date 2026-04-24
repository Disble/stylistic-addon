# Tasks: Remove Non-Replace Resolution Path

## Phase 1: Contract Lock-In

- [x] 1.1 Remove or rewrite tests in `src/adapters/word/WordAdapterAcceptSuggestion.test.ts` that model non-replace tracked-change recovery.
- [x] 1.2 Remove or rewrite tests in `src/adapters/word/WordAdapterRejectSuggestion.test.ts` that model non-replace tracked-change recovery, if any exist.
- [x] 1.3 Add or retain one focused contract test that invalid `track-change` suggestions fail fast instead of entering recovery.

## Phase 2: Core Refactor

- [x] 2.1 Modify `src/adapters/word/ResolveSuggestionCommand.ts` so `track-change` resolution validates the replace contract and never enters a non-replace execution branch.
- [x] 2.2 Modify `src/adapters/word/resolution/SuggestionResolutionObserver.ts` so tracked-change observation is replace-only.
- [x] 2.3 Modify `src/adapters/word/resolution/ResolutionSnapshotObserver.ts` so snapshot capture reflects only comment-only vs replace semantics.

## Phase 3: Cleanup

- [x] 3.1 Delete helper methods and logs that only supported the removed non-replace branch.
- [x] 3.2 Keep JSDoc aligned with the stricter tracked-change contract.

## Phase 4: Verification

- [x] 4.1 Run focused tests for `src/adapters/word/WordAdapterAcceptSuggestion.test.ts` and `src/adapters/word/WordAdapterRejectSuggestion.test.ts`.
- [x] 4.2 Run targeted Problems/typecheck validation for touched files.
- [x] 4.3 Re-read the final workflow and confirm tracked-change resolution has no non-replace operational branch.
