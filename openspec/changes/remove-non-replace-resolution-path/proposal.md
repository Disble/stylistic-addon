# Proposal: Remove Non-Replace Resolution Path

## Intent

Refactor suggestion resolution so every valid `track-change` suggestion is treated as a replace workflow and any `track-change` payload that cannot satisfy that contract fails fast as invalid input. Remove the adapter branch and tests that currently model a non-replace tracked-change resolution path.

## Scope

### In Scope
- Remove the non-replace tracked-change execution branch from `ResolveSuggestionCommand`.
- Treat invalid `track-change` suggestions as contract violations with explicit fail-fast behavior.
- Delete focused tests that protect the removed non-replace branch.
- Keep comment-only handling unchanged.

### Out of Scope
- Changing backend suggestion contracts beyond enforcing the existing one.
- Reworking replace observer evidence sources.
- UI redesign beyond surfacing the returned adapter error.

## Approach

Promote the existing contract to an invariant at resolution time:

1. `comment-only` suggestions keep their dedicated branch.
2. every valid `track-change` suggestion SHALL be resolved as replace.
3. a `track-change` suggestion without non-empty `anchor` and `suggestedText` SHALL be rejected immediately as invalid.
4. tests that modeled non-replace tracked-change recovery SHALL be deleted.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/adapters/word/resolve-suggestion/ResolveSuggestionCommand.ts` | Modified | Remove the non-replace execution branch and enforce a track-change invariant |
| `src/adapters/word/resolve-suggestion/SuggestionResolutionObserver.ts` | Modified | Remove non-replace observation normalization for tracked-change resolution |
| `src/adapters/word/resolve-suggestion/ResolutionSnapshotObserver.ts` | Modified | Treat tracked-change suggestions as replace snapshots by contract |
| `src/adapters/word/__tests__/WordAdapterAcceptSuggestion.test.ts` | Modified | Delete tests that only defend non-replace tracked-change recovery |
| `src/adapters/word/__tests__/WordAdapterRejectSuggestion.test.ts` | Modified | Delete non-replace-specific coverage if present |
| `openspec/changes/remove-non-replace-resolution-path/*` | New | SDD artifacts for the contract-hardening refactor |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Real Word still exposes a valid tracked-change case that is not a replace pair | Low | The backend contract and apply path already reject that shape; fail-fast preserves correctness over silent fallback |
| Tests currently rely on impossible fixtures | High | Delete only fixtures that violate the contract and keep replace/comment-only coverage intact |

## Rollback Plan

Revert the refactor and restore the removed branch/tests if real host evidence proves a valid product case still needs tracked-change resolution outside replace semantics.

## Dependencies

- `src/adapters/mastra/MastraAdapter.ts` suggestion validation contract
- `src/adapters/word/ApplySuggestionCommand.ts` tracked-change insertion contract
- Current resolution workflow in `src/adapters/word/resolve-suggestion/ResolveSuggestionCommand.ts`

## Success Criteria

- [ ] `ResolveSuggestionCommand` no longer contains a tracked-change non-replace execution branch.
- [ ] invalid `track-change` suggestions fail fast with an explicit contract error.
- [ ] branch-specific non-replace tests are removed.
- [ ] focused accept/reject validation remains green for supported workflows.
