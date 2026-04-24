# Design: Remove Non-Replace Resolution Path

## Technical Approach

The adapter already receives only two domain suggestion modes: `comment-only` and `track-change`. Because the backend contract and apply contract already require non-empty `anchor` and `suggestedText` for valid `track-change` suggestions, resolution should stop inferring an extra operational category. The command will therefore branch only by domain mode:

- `comment-only` -> comment-only cleanup flow
- `track-change` -> replace workflow

Any `track-change` value that cannot satisfy replace preconditions becomes invalid input and returns an error without attempting recovery.

## Architecture Decisions

| Decision | Choice | Alternatives | Rationale |
|---|---|---|---|
| Resolution mode split | Branch by domain type only | Keep a heuristic `replace vs non-replace` split inside resolution | The contract already defines the valid tracked-change shape, so the heuristic only hides invalid inputs |
| Invalid tracked-change handling | Fail fast with explicit error result | Retry with local fallback logic | Missing `anchor` or `suggestedText` is a contract violation, not a recoverable host-shape variation |
| Observer behavior | Treat tracked-change observation as replace-only | Preserve generic non-replace tracked-change observation | The retained product workflow only needs replace observation semantics |
| Snapshot behavior | Snapshot only comment-only vs replace | Preserve a third tracked-change/non-replace interpretation | Snapshots should describe supported workflows, not impossible intermediate categories |

## Data Flow

```text
Resolve suggestion
  -> if comment-only: cleanup comment and CC
  -> else validate tracked-change contract
       -> invalid: return fail-fast error
       -> valid: run replace workflow
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/adapters/word/ResolveSuggestionCommand.ts` | Modify | Replace `isReplaceSuggestion` heuristic with tracked-change contract assertion and remove non-replace execution path |
| `src/adapters/word/resolution/SuggestionResolutionObserver.ts` | Modify | Remove non-replace tracked-change observation branch |
| `src/adapters/word/resolution/ResolutionSnapshotObserver.ts` | Modify | Derive replace snapshot mode directly from suggestion type/contract |
| `src/adapters/word/WordAdapterAcceptSuggestion.test.ts` | Modify | Delete impossible non-replace contract tests |
| `src/adapters/word/WordAdapterRejectSuggestion.test.ts` | Modify | Delete impossible non-replace contract tests if any exist |

## Interfaces / Contracts

- `comment-only` SHALL remain the only non-tracked-change resolution mode.
- `track-change` SHALL require non-empty `anchor` and `suggestedText` throughout resolution.
- Resolution SHALL return an error result instead of attempting recovery when that invariant is broken.

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Adapter | Supported accept/reject tracked-change flows | Keep existing replace-focused tests green |
| Adapter | Invalid tracked-change invariant | Add/retain one focused fail-fast test if needed; delete impossible recovery tests |
| Adapter | Comment-only behavior | Keep existing coverage unchanged |

## Migration / Rollout

No migration required. This aligns runtime behavior with the contract already enforced upstream.

## Open Questions

- Should the fail-fast error message mention backend contract violation explicitly or keep adapter wording generic? Default: explicit contract wording.
