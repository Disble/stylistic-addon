# Design: Extract Replace Resolution Strategy

## Technical Approach

Refactor only the policy seam, not the workflow. `ResolveSuggestionCommand` remains the semantic orchestrator for locate -> observe -> execute -> reobserve -> cleanup. A new replace strategy object centralizes the action-dependent decisions that are currently duplicated across the command and the executor.

## Architecture Decisions

| Decision | Choice | Alternatives | Rationale |
|---|---|---|---|
| Strategy scope | Policy-only strategy | Separate accept/reject command classes | The workflow is shared; only a few replace decisions vary |
| Strategy location | `src/adapters/word/resolution/` | Keep helper methods inside command | The executor also needs the same policy |
| Wiring | Build once in `ResolveSuggestionCommand` and pass to executor | Let each class infer from `action` independently | One source of truth avoids behavioral drift |

## Data Flow

```text
ResolveSuggestionCommand
  -> builds ReplaceResolutionStrategy from action
  -> uses strategy.semanticOrder() for replace steps
  -> passes strategy to TrackedChangeResolutionExecutor
  -> executor uses strategy.priorityFor(type)
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/adapters/word/resolution/ReplaceResolutionStrategyContext.ts` | Create | Define the policy contract and concrete accept/reject strategy creation |
| `src/adapters/word/ResolveSuggestionCommand.ts` | Modify | Replace inline order/label branching with strategy usage |
| `src/adapters/word/resolution/TrackedChangeResolutionExecutor.ts` | Modify | Consume strategy priority instead of raw action checks |
| `src/adapters/word/WordAdapterAcceptSuggestion.test.ts` | Modify | Keep accept ordering/contract assertions stable through refactor |
| `src/adapters/word/WordAdapterRejectSuggestion.test.ts` | Modify | Keep reject ordering/contract assertions stable through refactor |

## Interfaces / Contracts

```ts
export type ReplaceTrackedChangeSide = "Added" | "Deleted";

export interface ReplaceResolutionStrategy {
  readonly actionLabel: "aceptación" | "rechazo";
  readonly semanticOrder: readonly [ReplaceTrackedChangeSide, ReplaceTrackedChangeSide];
  priorityFor(type: string): number;
}
```

`priorityFor()` stays tolerant of non-replace types by returning a fallback priority.

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Adapter | Accept replace still runs `Added -> Deleted` | Keep existing focused accept tests green |
| Adapter | Reject replace still runs `Deleted -> Added` | Keep existing focused reject tests green |
| Unit | Strategy policy contract | Add a small direct test if extraction logic becomes non-trivial |
| Integration | Command and executor share one policy source | Assert no behavioral drift through focused adapter suites |

## Migration / Rollout

No migration required.

## Open Questions

- [ ] Whether `actionLabel` should live in the strategy or in a smaller formatter helper.
