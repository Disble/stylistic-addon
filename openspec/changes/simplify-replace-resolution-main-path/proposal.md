# Proposal: Simplify Replace Resolution Main Path

## Intent

Refactor replace suggestion resolution so the code keeps only one primary orchestration strategy: semantic stepwise execution with fresh re-observation and a single bounded atomic retry after post-execute verification. Remove earlier atomic fallback branches and delete tests that only protect those discarded branches.

## Scope

### In Scope
- Keep `accept`/`reject` replace resolution centered on semantic side execution with fresh proxies.
- Keep the fail-closed post-execute verification gate.
- Keep at most one bounded atomic retry after a fresh post-execute observation still exposes the full replace pair.
- Remove early atomic fallback branches that overlap with the surviving post-execute retry.
- Remove tests that validate only deleted fallback paths instead of the retained workflow contract.

### Out of Scope
- Reworking observation identity or locator strategy beyond what the retained workflow already consumes.
- UI/state-machine redesign outside the adapter contract affected by the refactor.
- Broad cleanup of unrelated accept/reject test scenarios.

## Approach

Promote this orchestration as the only supported replace-resolution workflow:

1. observe a confirmed pending replace pair,
2. normalize duplicate semantic sides,
3. execute the semantic primary side with fresh verification,
4. re-observe only the remaining side,
5. execute the remaining side,
6. verify post-execute state,
7. if a fresh post-execute observation still exposes the full pair, try one bounded atomic retry,
8. if that retry fails or still cannot certify completion, use one fresh semantic recovery pass and otherwise fail closed.

The key change is deleting the earlier atomic escape hatches that run before the stronger post-execute verification point. Those branches increase complexity, overlap with the surviving retry, and force dedicated tests that no longer describe the desired main workflow.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/adapters/word/ResolveSuggestionCommand.ts` | Modified | Remove early atomic fallback helpers and simplify replace execution orchestration. |
| `src/adapters/word/WordAdapterAcceptSuggestion.test.ts` | Modified | Keep only tests that protect the retained replace-resolution workflow. |
| `openspec/changes/simplify-replace-resolution-main-path/*` | New | Proposal, spec, design, and tasks for the refactor. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| A removed fallback covered a real host shape not exercised by the retained post-execute retry | Medium | Keep focused regression tests for stepwise execution, fail-closed verification, and the surviving post-execute atomic retry. |
| Test deletion removes useful unrelated signal | Medium | Delete only tests tied to removed helper branches; keep tests that protect the retained public contract. |

## Rollback Plan

Revert the `ResolveSuggestionCommand` refactor and restore the deleted fallback-specific tests if real Word proves the removed early atomic paths were still required.

## Dependencies

- Existing replace diagnostics in `docs/replace-resolution-postmortem.md`.
- Current replace workflow implementation in `src/adapters/word/ResolveSuggestionCommand.ts`.

## Success Criteria

- [ ] `ResolveSuggestionCommand` no longer contains early atomic fallback helpers that precede the post-execute verification gate.
- [ ] Replace resolution keeps one primary semantic-step workflow plus one bounded post-execute atomic retry.
- [ ] Tests that only defend removed fallback branches are deleted.
- [ ] Focused validation still proves the retained workflow contract.
