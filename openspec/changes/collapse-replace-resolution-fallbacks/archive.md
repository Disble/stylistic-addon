# Archive: Collapse Replace Resolution Fallbacks

## Outcome

Completed. `ResolveSuggestionCommand` now uses one semantic execution model plus one shared fresh-proxy retry policy.

## Implemented

- Removed the post-execute atomic accept fallback.
- Removed the semantic recovery path that only existed to rescue that atomic fallback.
- Removed replace silent-no-op body-text recovery and its text-matching helper.
- Renamed the non-replace recovery path to the shared fresh-proxy retry policy and kept it bounded to one retry.
- Rebuilt the focused accept/reject adapter suites so they defend only the retained workflow contract.

## Validation

- Focused adapter tests: 11 passed, 0 failed.
- Problems check: clean for `src/adapters/word/ResolveSuggestionCommand.ts`, `src/adapters/word/WordAdapterAcceptSuggestion.test.ts`, and `src/adapters/word/WordAdapterRejectSuggestion.test.ts`.

## Residual Risk

Real Word host behavior may still expose new observer-edge cases, but the retained behavior now fails closed instead of switching to alternate execution modes.
