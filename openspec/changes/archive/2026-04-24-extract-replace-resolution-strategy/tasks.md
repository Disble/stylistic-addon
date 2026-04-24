# Tasks: Extract Replace Resolution Strategy

## Phase 1: Foundation

- [x] 1.1 Create `src/adapters/word/resolution/ReplaceResolutionStrategyContext.ts` with the shared replace policy contract and accept/reject implementations.
- [x] 1.2 Define the minimal exported types for semantic sides and strategy creation without moving workflow code into the new file.

## Phase 2: Core Refactor

- [x] 2.1 Modify `src/adapters/word/ResolveSuggestionCommand.ts` to build/use the shared strategy for semantic order and action labels.
- [x] 2.2 Modify `src/adapters/word/resolution/TrackedChangeResolutionExecutor.ts` to consume the shared strategy for tracked-change priority ordering.
- [x] 2.3 Remove now-redundant inline action-branch helpers once both call sites use the strategy.

## Phase 3: Verification

- [x] 3.1 Update `src/adapters/word/WordAdapterAcceptSuggestion.test.ts` only where needed to keep accept ordering assertions stable.
- [x] 3.2 Update `src/adapters/word/WordAdapterRejectSuggestion.test.ts` only where needed to keep reject ordering assertions stable.
- [x] 3.3 Run focused adapter validation for the accept/reject suites and a Problems check for the touched files.

## Phase 4: Cleanup

- [x] 4.1 Align JSDoc/comments in `ResolveSuggestionCommand.ts` and `TrackedChangeResolutionExecutor.ts` with the new shared-policy model.
- [x] 4.2 Archive the change only after behavior remains unchanged and duplication is removed.
