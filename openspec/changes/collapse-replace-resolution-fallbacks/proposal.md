# Proposal: Collapse Replace Resolution Fallbacks

## Intent

Refactor replace suggestion resolution so the adapter keeps one semantic main path and one shared fresh-proxy retry policy. Remove fallback branches that currently duplicate or weaken that path: non-replace same-click recovery, replace body-text silent-no-op recovery, post-execute atomic retry, and the semantic retry path that exists only to recover from that atomic retry.

## Scope

### In Scope
- Keep replace resolution centered on semantic side execution with fresh re-observation.
- Keep one bounded retry model: re-locate, re-observe, retry once with fresh proxies, then fail closed.
- Remove fallback branches that bypass or duplicate the semantic workflow.
- Update focused adapter tests so they defend only the retained workflow contract.

### Out of Scope
- Rewriting locator or observer evidence-source strategy.
- Taskpane/UI state-machine changes outside adapter results.
- Broad cleanup of unrelated comment-only or telemetry behavior.

## Approach

Promote this as the only supported resolution model:

1. locate and observe the suggestion,
2. execute `apply` on the expected tracked change(s),
3. re-observe fresh Word state to certify progress,
4. retry once with fresh proxies when certification fails,
5. fail closed if fresh observation still cannot certify completion.

For replace, the semantic order remains action-dependent: `Added -> Deleted` for accept and `Deleted -> Added` for reject.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/adapters/word/ResolveSuggestionCommand.ts` | Modified | Remove redundant fallback branches and consolidate recovery around one semantic retry model |
| `src/adapters/word/WordAdapterAcceptSuggestion.test.ts` | Modified | Replace fallback-specific tests with retained-contract coverage |
| `src/adapters/word/WordAdapterRejectSuggestion.test.ts` | Modified | Replace fallback-specific tests with retained-contract coverage |
| `openspec/changes/collapse-replace-resolution-fallbacks/*` | New | Proposal, spec, design, and tasks for the refactor |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| A removed fallback still covers a real Word host shape | Medium | Keep focused RED/GREEN coverage for semantic ordering, fresh-proxy retry, and fail-closed certification |
| Collapsing recoveries hides an evidence-source gap in the observer | Medium | Keep the observer contract unchanged in this change and fail closed when re-observation cannot certify completion |

## Rollback Plan

Revert the `ResolveSuggestionCommand` refactor and restore the deleted fallback-specific tests if real Word proves the removed branches still cover required host behavior.

## Dependencies

- Current workflow in `src/adapters/word/ResolveSuggestionCommand.ts`
- Existing focused adapter suites for accept/reject replace resolution

## Success Criteria

- [x] `ResolveSuggestionCommand` exposes one semantic resolution path plus one shared fresh-proxy retry policy.
- [x] Post-execute atomic retry and its follow-up semantic recovery are removed.
- [x] Body-text silent-no-op recovery and non-replace same-click recovery are removed.
- [x] Focused adapter validation still protects the retained contract and fails closed on uncertified resolution.
