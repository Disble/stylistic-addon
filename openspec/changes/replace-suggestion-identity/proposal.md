# Proposal: Replace Suggestion Identity

## Intent

Resolve recurrent `already-resolved` regressions for `replace` suggestions in Word. The system incorrectly promotes observability failures into terminal resolution because it confuses a partial operational handle (Content Control on the inserted side) with the full domain identity of a replace suggestion.

## Scope

### In Scope
- Define a compound identity model for replace suggestions (domain identity + multiple Word artifact references).
- Defensive phase: Stop mapping `0 tracked changes observed` to `already-resolved`. Introduce non-terminal states like `unobservable` or `identity-lost`.
- Structural phase: Implement the new domain identity for replace suggestions.

### Out of Scope
- UI redesign for the taskpane.
- Changing identity models for non-replace suggestions (`insert` or `delete`), unless required by the new generic model.

## Approach

**Phased approach:**

1. **Defensive Phase:** Update `WordAdapter` and resolution logic to return a non-terminal state (`unobservable` or `identity-lost`) instead of `already-resolved` when tracked changes are not found. This immediately stops false positives in the taskpane and avoids sending incorrect feedback.
2. **Structural Phase:** Evolve the domain to use a compound identity model (`ReplaceSuggestionIdentity`). Treat a replace suggestion as a composed review unit distinguishing the inserted-side reference, deleted-side reference, and operational anchors. Update `ApplySuggestionCommand` and `WordAdapter` to separate locating Word refs, observing semantic replace state, resolving tracked changes, and concluding business status.

**Key Rule:** NEVER upgrade observability failure into terminal resolution.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `WordAdapter` | Modified | Update resolution logic to avoid false `already-resolved` and separate observation from business status. |
| `ApplySuggestionCommand` | Modified | Stop treating a single inserted-side Content Control as the entire identity. |
| Domain Models | Modified/New | Introduce `ReplaceSuggestionIdentity`, `ObservationStatus`, and `WordArtifactRef`. |
| Taskpane/Workflow | Modified | Handle new non-terminal states (`unobservable`, `identity-lost`) appropriately. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Taskpane drops into unhandled state for `unobservable` status | Medium | Add explicit handlers in UI and workflow state machines for new non-terminal states before returning them from the adapter. |

## Rollback Plan

Revert the defensive phase mapping changes in `WordAdapter` to restore the previous `already-resolved` fallback. For the structural phase, maintain the legacy identity model behind a feature toggle until the compound identity is proven stable across all supported Word versions.

## Dependencies

- Existing domain documentation: `docs/replace-suggestion-identity-proposal.md` and `docs/review-domain-and-track-changes.md`.

## Success Criteria

- [ ] Replace suggestions with unobservable tracked changes no longer incorrectly mark as `already-resolved`.
- [ ] `WordAdapter` explicitly models and returns `unobservable` or `identity-lost` states.
- [ ] A compound identity model is implemented and utilized for replace suggestions, capturing both inserted and deleted sides.