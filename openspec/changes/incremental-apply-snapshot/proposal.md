# Proposal: Incremental Apply Snapshot

## Intent

Reduce `backend context stale` failures during batch apply by making each successful mutation in real Word return enough local state to update a live frontend snapshot and rebase pending suggestions against the new document state.

## Scope

### In Scope
- Define an incremental apply model based on `apply in Word -> patch local snapshot -> rebase pending suggestions`.
- Add contracts for apply mutation patches, paragraph-local rereads, and snapshot versioning.
- Introduce batch ordering/ranking seams that can use real document position or snapshot-local position instead of raw backend array order.

### Out of Scope
- Re-reading the full document after every applied suggestion.
- Backend-side semantic diffing or external storage of document versions.

## Approach

Use Word as the source of truth. After each real apply, capture a localized patch (affected paragraph/range, delta length, updated text) and update a live snapshot in the batch workflow. Rebase pending suggestions incrementally, with localized rereads for hot paragraphs and existing text-location heuristics reserved as fallback.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/domain/types.ts` | Modified | Add snapshot, patch, and rebase contracts for apply batches. |
| `src/adapters/word/ApplySuggestionCommand.ts` | Modified | Return mutation patch/state from real Word apply operations. |
| `src/adapters/word/BatchApplyOrchestrator.ts` | Modified | Maintain live snapshot, rebase pending suggestions, and replace fake reverse ordering. |
| `src/adapters/word/WordAdapter.ts` | Modified | Provide initial text/snapshot inputs and wire the new batch behavior. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Snapshot drifts from Word after local patching | Medium | Re-read affected paragraphs/ranges after each successful apply and keep Word authoritative. |
| Rebase logic becomes too heavy for large documents | Medium | Limit rereads to local hot zones; keep global fallback heuristics only for misses. |

## Rollback Plan

Revert batch apply to the current per-suggestion isolated execution path with heuristic re-location and remove the new snapshot/rebase contracts if the incremental model proves unstable.

## Dependencies

- Existing apply search heuristics in `src/core/text-search/TextSearchCore.ts` and `src/adapters/word/WordTextLocatorAdapter.ts`.
- Existing batch apply flow in `src/adapters/word/BatchApplyOrchestrator.ts`.

## Success Criteria

- [ ] Batch apply can rebase pending suggestions against a live snapshot without re-reading the entire document.
- [ ] Hot-paragraph mutations no longer depend solely on stale backend contexts.
- [ ] Batch ordering stops pretending reverse array order equals real document position.
