# Design: Semantic Reconciliation Telemetry

## Technical Approach

The current resolution workflow assumes one atomic success path: observe → execute → cleanup → inspect → map result. Real Word behavior breaks that assumption because tracked changes may mutate successfully before later ranges/comments become invalid. The new design adds semantic reconciliation and structured telemetry so the workflow can distinguish **semantic outcome** from **housekeeping visibility**.

Target flow:

```text
observe-before
  -> execute tracked changes (report attempted/completed/failure)
  -> reconcile semantic state if any phase fails or host state becomes suspect
  -> cleanup as best effort
  -> inspect-after
  -> map terminal status + warnings + telemetry context
```

## Architecture Decisions

| Decision | Choice | Alternatives | Rationale |
|---|---|---|---|
| Source of truth after failures | Re-observe Word semantics, not thrown exceptions | Treat every exception as retryable error | Word is non-atomic; exceptions do not prove semantic failure |
| Executor contract | Return `ResolutionExecutionReport` instead of `void` | Infer progress from logs/side effects | Guardrails need explicit evidence of what already executed |
| Telemetry architecture | Add non-blocking `ITelemetryPort` with console-first adapter | Keep ad-hoc `console.log` only | Structured events enable future sinks without coupling workflow semantics |
| Cleanup semantics | Accept and reject both use best-effort cleanup with warnings | Keep accept rethrow asymmetry | Cleanup truth is secondary to semantic resolution truth |

## Data Flow

```text
SuggestionCardRenderer
  -> SuggestionResolutionWorkflow
    -> WordAdapter.resolveSuggestion(action)
      -> ResolveSuggestionCommand
        -> SuggestionLocator
        -> SuggestionResolutionObserver
        -> TrackedChangeResolutionExecutor
        -> SuggestionResolutionReconciler (new)
        -> SuggestionResolutionCleanup
        -> ResolveSuggestionResultFactory
      -> ITelemetryPort.emit(event)
```

Telemetry emits phase facts; reconciliation decides semantic status; result factory merges both into one adapter result.

## Invariants

- Semantic success MUST out-rank cleanup or observability failure.
- Telemetry MUST be best effort and MUST NOT mutate workflow semantics.
- Retryable `error` MUST mean semantic state is still unresolved or unknown after reconciliation.
- Terminal results MAY carry warnings, but warnings MUST NOT resurrect actions in the taskpane.

## Interfaces / Contracts

```ts
type ResolutionPhase =
  | "observe-before"
  | "execute"
  | "reconcile"
  | "cleanup"
  | "inspect-after";

interface ResolutionExecutionReport {
  attempted: number;
  completed: number;
  failureIndex?: number;
  failureMessage?: string;
}

interface SuggestionResolutionWarning {
  kind: "cleanup-failed" | "inspection-failed" | "telemetry-failed";
  message: string;
  phase: ResolutionPhase;
}

interface ResolutionTelemetryEvent {
  workflowAttemptId: string;
  suggestionId: string;
  action: "accept" | "reject";
  phase: ResolutionPhase;
  outcome: "started" | "completed" | "failed" | "reconciled";
  metadata?: Record<string, string | number | boolean>;
}
```

## File Changes

| File | Action | Description |
|---|---|---|
| `src/domain/types.ts` | Modify | Add execution-report, warning, and telemetry event/result metadata types. |
| `src/domain/ports.ts` | Modify | Add `ITelemetryPort` as best-effort observability contract. |
| `src/adapters/word/resolution/TrackedChangeResolutionExecutor.ts` | Modify | Return `ResolutionExecutionReport`. |
| `src/adapters/word/resolution/SuggestionResolutionReconciler.ts` | Create | Re-observe after failures and classify semantic outcome. |
| `src/adapters/word/resolution/SuggestionResolutionCleanup.ts` | Modify | Return warnings instead of rethrowing late cleanup failures. |
| `src/adapters/word/ResolveSuggestionCommand.ts` | Modify | Orchestrate report, reconciliation, telemetry, and warning-aware result mapping. |
| `src/adapters/telemetry/ConsoleTelemetryAdapter.ts` | Create | Default sink that logs structured events without affecting behavior. |
| `src/domain/suggestion/SuggestionResolutionWorkflow.ts` | Modify | Preserve terminal status + warnings and emit feedback only from semantic outcome. |
| `src/domain/suggestion/SuggestionStateMachine.ts`, `src/taskpane/SuggestionCardRenderer.ts` | Modify | Support terminal-warning UI semantics. |
| `docs/TELEMETRY.md` | Create | Document observability rules and naming conventions. |

## Testing Strategy

| Layer | What to Test | Approach |
|---|---|---|
| Tier 1 adapter | Partial accept/reject, late cleanup failure, semantic reconciliation | `WordAdapterAcceptSuggestion.test.ts`, `WordAdapterRejectSuggestion.test.ts` |
| Tier 1 resolution modules | execution reports, reconciler classification, telemetry swallowing | focused resolution collaborator tests |
| Tier 2 taskpane | terminal warning state, no action resurrection | `TaskpaneSuggestionResolution.test.ts` |
| Docs/policy | telemetry naming and invariants | `docs/TELEMETRY.md` review during implementation |

## Migration / Rollout

No backend migration required. Roll out in three slices: (1) contracts + RED regressions, (2) reconciliation semantics in adapter/workflow, (3) console telemetry + docs. Remote sinks stay future work.

## Risks / Tradeoffs / Open Questions

- Tradeoff: richer result contracts add test surface, but they remove semantic lies.
- Risk: ambiguous post-failure observation may still end as retryable; this is acceptable if explicit.
- Open question: whether warning copy belongs in the adapter result or only in taskpane formatting helpers.

## Recommended Implementation Sequence

1. Add RED regressions for partial accept/reject and terminal-warning UI.
2. Introduce execution report + reconciler.
3. Make cleanup warning-based on both actions.
4. Add telemetry port/console adapter and emit structured phase events.
5. Document conventions in `docs/TELEMETRY.md`.
