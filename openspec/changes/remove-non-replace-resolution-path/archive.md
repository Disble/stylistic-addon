# Archive: Remove Non-Replace Resolution Path

## Outcome

Completed. Tracked-change resolution now has only two supported outcomes at the adapter boundary: valid replace workflow or explicit invalid-contract failure.

## Implemented

- Removed the tracked-change non-replace execution branch from `ResolveSuggestionCommand`.
- Added fail-fast contract enforcement for `track-change` suggestions missing `anchor` or `suggestedText`.
- Simplified `SuggestionResolutionObserver` to use replace observation only for tracked-change resolution.
- Simplified `ResolutionSnapshotObserver` to distinguish only comment-only vs tracked-change snapshots.
- Replaced the impossible non-replace accept test with a fail-fast contract test.

## Validation

- Focused adapter tests: 11 passed, 0 failed.
- Problems check: clean for `src/adapters/word/resolve-suggestion/ResolveSuggestionCommand.ts`, `src/adapters/word/resolve-suggestion/SuggestionResolutionObserver.ts`, and `src/adapters/word/__tests__/WordAdapterAcceptSuggestion.test.ts`.
- Code search: no remaining `non-replace` or `isReplaceSuggestion()` branch markers in `src/adapters/word/**`.

## Residual Risk

If real Word host behavior ever produces a legitimate tracked-change case outside replace semantics, the adapter will now fail fast instead of attempting silent recovery. That is intentional because the current backend/apply contract defines such input as invalid.
