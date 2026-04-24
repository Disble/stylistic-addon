# Archive: Extract Replace Resolution Strategy

## Outcome

Completed. Replace semantic policy now comes from one shared strategy used by both `ResolveSuggestionCommand` and `TrackedChangeResolutionExecutor`.

## Implemented

- Added `ReplaceResolutionStrategyContext.ts` as the shared source for action label, semantic order, and replace-step priority.
- Updated `ResolveSuggestionCommand` to consume the shared replace strategy.
- Updated `TrackedChangeResolutionExecutor` to consume the same shared replace strategy.
- Added focused strategy coverage and kept existing accept/reject behavior green.

## Validation

- Focused tests: 20 passed, 0 failed.
- Problems check: clean for the touched command, executor, strategy, and focused accept/reject suites.

## Residual Risk

The strategy extraction is behavioral-preserving, so the remaining risk is limited to future Word host edge cases already covered by the retained workflow and focused adapter suites.
