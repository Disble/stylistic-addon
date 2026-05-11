# Tasks: Collapse Replace Resolution Fallbacks

## Phase 1: Contract Lock-In

- [x] 1.1 Add RED tests in `src/adapters/word/__tests__/WordAdapterAcceptSuggestion.test.ts` for the retained accept contract: semantic `Added -> Deleted`, fresh re-observation, fail-closed certification without atomic fallback, and one bounded non-replace fresh-proxy retry contract.
- [x] 1.2 Add RED tests in `src/adapters/word/__tests__/WordAdapterRejectSuggestion.test.ts` for the retained reject contract: semantic `Deleted -> Added`, fresh re-observation, and fail-closed certification without body-text recovery.
- [x] 1.3 Delete or rewrite tests that only prove `applyAtomically`, body-text silent-no-op recovery, or non-replace same-click fallback behavior.

## Phase 2: Core Refactor

- [x] 2.1 Modify `src/adapters/word/resolve-suggestion/ResolveSuggestionCommand.ts` so non-replace resolution becomes `apply -> reobserve -> retry once -> reobserve -> certify/fail`.
- [x] 2.2 Modify `src/adapters/word/resolve-suggestion/ResolveSuggestionCommand.ts` so replace resolution keeps only semantic side execution plus one shared fresh-proxy retry policy.
- [x] 2.3 Remove `applyAtomically`-based post-execute recovery and delete the semantic recovery path that exists only to rescue that branch.
- [x] 2.4 Remove body-text silent-no-op recovery and any helper logic that resolves tracked changes by text matching instead of suggestion identity.

## Phase 3: Cleanup

- [x] 3.1 Delete unused helper methods, logging branches, and report-merging paths in `src/adapters/word/resolve-suggestion/ResolveSuggestionCommand.ts` after the fallback collapse.
- [x] 3.2 Keep JSDoc and inline intent comments aligned with the new single-workflow model.

## Phase 4: Verification

- [x] 4.1 Run focused RED/GREEN validation for `src/adapters/word/__tests__/WordAdapterAcceptSuggestion.test.ts` and `src/adapters/word/__tests__/WordAdapterRejectSuggestion.test.ts`.
- [x] 4.2 Run targeted Problems/typecheck validation for `src/adapters/word/resolve-suggestion/ResolveSuggestionCommand.ts` and any touched helpers.
- [x] 4.3 Re-read the final workflow and confirm the code exposes only one semantic execution model plus one shared fresh-proxy retry policy.
