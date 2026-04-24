# Proposal: Extract Replace Resolution Strategy

## Intent

Extract the `accept` vs `reject` replace-policy decisions into one small strategy so `ResolveSuggestionCommand` keeps a single workflow skeleton and stops duplicating semantic-order logic, action labels, and step priority rules.

## Scope

### In Scope
- Introduce one strategy contract for replace semantic policy.
- Move action-specific replace ordering and labels out of `ResolveSuggestionCommand`.
- Reuse the same policy in `TrackedChangeResolutionExecutor`.
- Update focused tests to protect behavior while allowing the refactor.

### Out of Scope
- Changing the observed replace workflow or retry policy.
- Rewriting locator, observer, cleanup, or taskpane behavior.
- Generalizing non-replace resolution behind new abstractions.

## Approach

Keep `ResolveSuggestionCommand` as the template/orchestrator. Add a small strategy object that answers: semantic order, primary side, secondary side, and action label. Inject or build it once per command and pass it to the executor so both layers read the same policy source.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/adapters/word/ResolveSuggestionCommand.ts` | Modified | Replace inline `accept`/`reject` branching with strategy usage |
| `src/adapters/word/resolution/TrackedChangeResolutionExecutor.ts` | Modified | Read step priority from the shared strategy instead of duplicating action checks |
| `src/adapters/word/resolution/*` | New | Add the replace strategy contract/implementation |
| `src/adapters/word/WordAdapterAcceptSuggestion.test.ts` | Modified | Keep accept behavior coverage after extraction |
| `src/adapters/word/WordAdapterRejectSuggestion.test.ts` | Modified | Keep reject behavior coverage after extraction |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Strategy grows into a second workflow abstraction | Medium | Keep the contract limited to policy only: order and labels |
| Command/executor diverge during migration | Medium | Switch both call sites in the same change and keep focused accept/reject tests green |

## Rollback Plan

Revert the strategy extraction commit and restore the direct inline action branching in the command and executor.

## Dependencies

- Current replace workflow in `src/adapters/word/ResolveSuggestionCommand.ts`
- Current ordering logic in `src/adapters/word/resolution/TrackedChangeResolutionExecutor.ts`

## Success Criteria

- [ ] One shared strategy defines replace semantic policy for `accept` and `reject`.
- [ ] `ResolveSuggestionCommand` no longer owns hardcoded action-order branching.
- [ ] `TrackedChangeResolutionExecutor` no longer duplicates replace priority branching.
- [ ] Focused accept/reject tests still prove the same external behavior.
