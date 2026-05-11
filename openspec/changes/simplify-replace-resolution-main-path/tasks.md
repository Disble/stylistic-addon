# Tasks: Simplify Replace Resolution Main Path

## Phase 1: SDD guardrails

- [x] 1.1 Add a spec for the retained replace-resolution workflow and explicitly declare early atomic fallback branches out of scope.
- [x] 1.2 Identify the production helpers and tests that exist only to support the removed early atomic paths.

## Phase 2: Production refactor

- [x] 2.1 Modify `src/adapters/word/resolve-suggestion/ResolveSuggestionCommand.ts` to remove `tryAtomicAcceptReplaceFallback(...)`.
- [x] 2.2 Modify `src/adapters/word/resolve-suggestion/ResolveSuggestionCommand.ts` to remove `tryAtomicAcceptStepFallback(...)` and simplify callers.
- [x] 2.3 Keep the semantic stepwise workflow plus the bounded post-execute atomic retry and fresh semantic recovery.

## Phase 3: Test pruning

- [x] 3.1 Prune `src/adapters/word/__tests__/WordAdapterAcceptSuggestion.test.ts` to tests that directly protect the retained replace workflow contract.
- [x] 3.2 Prune `src/adapters/word/__tests__/WordAdapterRejectSuggestion.test.ts` to the symmetric retained replace workflow contract.

## Phase 4: Validation

- [x] 4.1 Run focused adapter tests for the retained replace-resolution workflow after pruning.
- [x] 4.2 Run a narrow typecheck or lint validation for touched files if needed.
