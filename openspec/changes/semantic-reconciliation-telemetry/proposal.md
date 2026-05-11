# Proposal: Semantic Reconciliation Telemetry

## Intent

Fix the `ItemNotFound` resolution lie where Word mutates a suggestion semantically, but the add-in still reports retryable `error`, leaves the card pending, and loses cleanup visibility. Add first-class telemetry so future regressions can be debugged from structured evidence instead of scattered console traces.

## Scope

### In Scope
- Add semantic reconciliation for accept/reject after execution, cleanup, or inspection failures.
- Add structured resolution telemetry with correlation ids, phase events, and warning metadata.
- Update taskpane/workflow semantics so terminal semantic success stays terminal even when housekeeping fails.
- Add `docs/TELEMETRY.md` as the observability guide for future changes.

### Out of Scope
- Backend telemetry ingestion, dashboards, or remote analytics.
- Re-architecting the whole analysis pipeline outside observability touchpoints.
- Broad UI redesign beyond terminal-warning rendering.

## Approach

Introduce a `ResolutionExecutionReport`, a `SuggestionResolutionReconciler`, and a non-blocking telemetry port. The workflow will become: observe before → execute tracked changes → reconcile semantic state if anything fails → run cleanup as best effort → emit terminal result plus warnings/telemetry metadata. Key rule: **never let housekeeping or telemetry failure overwrite semantic truth**.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/adapters/word/resolve-suggestion/` | Modified/New | Add execution report, reconciliation orchestration, cleanup warnings, telemetry emission. |
| `src/domain/types.ts`, `src/domain/ports.ts` | Modified | Add warning/telemetry contracts and telemetry port. |
| `src/domain/suggestion/`, `src/taskpane/` | Modified | Preserve terminal UI state after reconciled success; surface warnings. |
| `docs/TELEMETRY.md` | New | Establish observability principles and event conventions. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Over-modeling telemetry before implementation | Medium | Keep v1 sink console-backed and event schema minimal. |
| UI confusion between warning and error | Medium | Add explicit terminal-warning scenarios in workflow/taskpane tests. |
| Reconciliation misclassifies unresolved state | Medium | Add strict RED regressions for partial accept/reject and late invalidation paths. |

## Rollback Plan

Revert the reconciliation branch and telemetry contract additions, keeping the current resolution workflow. `docs/TELEMETRY.md` can remain as guidance because it is non-executable.

## Dependencies

- Existing resolution workflow docs in `docs/review-domain-and-track-changes.md`.
- Current resolution modules under `src/adapters/word/resolve-suggestion/`.

## Success Criteria

- [ ] Accept/reject no longer stay retryable when Word already resolved the suggestion semantically.
- [ ] Cleanup failures are exposed as warnings/telemetry, not as false semantic failure.
- [ ] Resolution attempts emit structured, correlated telemetry events.
- [ ] `docs/TELEMETRY.md` documents the observability mindset for future work.
