# Design: Incremental Apply Snapshot

## Technical Approach

Current batch apply executes suggestions sequentially in isolated `Word.run` contexts, but it keeps no live snapshot of the mutated Word text. That leaves later suggestions dependent on stale backend contexts and a fake ordering heuristic (`reverse()`). The new design keeps Word as the source of truth while maintaining a live incremental snapshot updated from each successful real apply.

Target flow:

```text
read initial text source
  -> build AnalysisSnapshot(version 0)
  -> rank suggestions by snapshot position
  -> apply suggestion in Word
  -> capture localized mutation patch from Word
  -> update snapshot(version + 1)
  -> rebase pending suggestions
  -> re-read only hot paragraphs when needed
  -> continue batch
```

## Architecture Decisions

| Decision | Choice | Alternatives | Rationale |
|---|---|---|---|
| Source of truth | Real Word mutation first, local snapshot second | Pure frontend simulation | Prevents drift between model and host document |
| Update granularity | Localized patch + hot-paragraph reread | Full-document reread after every apply | Large docs (e.g. 223k chars) make full rereads too expensive |
| Pending-suggestion update | Incremental rebase with offset deltas plus local rescoring | Re-run global heuristics for every pending suggestion | Keeps cost bounded and uses global locator only as fallback |
| Batch ordering | Explicit ranking seam from snapshot/local position | `reverse()` heuristic | Reverse array order is not real document order |

## Data Flow

```text
WordAdapter.getTextToAnalyze()
  -> BatchApplyOrchestrator.buildInitialSnapshot()
    -> ApplySuggestionCommand.execute()
      -> Word real mutation
      -> localized reread / mutation patch
    -> BatchSnapshotUpdater.applyPatch()
    -> PendingSuggestionRebaser.rebase()
    -> next suggestion
```

## Invariants

- Word MUST remain the source of truth for applied text.
- A local snapshot MUST only advance after a successful real Word mutation.
- Full-document rereads MUST NOT be required after every apply.
- Global text-location fallback MUST remain secondary to snapshot rebase + local reread.
- Batch ordering MUST come from real position signals, not backend array order alone.

## Interfaces / Contracts

```ts
interface AnalysisSnapshot {
  text: string;
  isSelection: boolean;
  version: number;
  paragraphs: SnapshotParagraph[];
}

interface SnapshotParagraph {
  paragraphId: string;
  text: string;
  startOffset: number;
  endOffset: number;
}

interface ApplyMutationPatch {
  suggestionId: string;
  snapshotVersion: number;
  paragraphId?: string;
  originalText: string;
  updatedText: string;
  deltaLength: number;
  affectedStart: number;
  affectedEnd: number;
}

interface RebasedSuggestion {
  suggestion: Suggestion;
  snapshotVersion: number;
  paragraphId?: string;
  anchorStart?: number;
  anchorEnd?: number;
  requiresLocalReread: boolean;
}
```

## File Changes

| File | Action | Description |
|---|---|---|
| `src/domain/types.ts` | Modify | Add snapshot, patch, and rebased-suggestion contracts |
| `src/adapters/word/ApplySuggestionCommand.ts` | Modify | Return `ApplyMutationPatch` / localized reread result from real Word apply |
| `src/adapters/word/BatchApplyOrchestrator.ts` | Modify | Maintain live snapshot, ranking seam, and incremental rebase |
| `src/adapters/word/WordAdapter.ts` | Modify | Supply initial text source and any required snapshot helpers |
| `src/adapters/word/BatchApplyOrchestrator.test.ts` | Create | Add direct Tier 1 coverage for ordering/rebase policy |

## Testing Strategy

| Layer | What to Test | Approach |
|---|---|---|
| Tier 1 batch | Misordered backend arrays, hot-paragraph rebase, local reread escalation | New `BatchApplyOrchestrator.test.ts` with strict doubles |
| Tier 1 apply | Successful apply returns localized patch from real Word mutation | `ApplySuggestionCommand*.test.ts` |
| Tier 1 adapter | Initial snapshot creation from selection/body | `WordAdapterReadText.test.ts`, `WordAdapterApplySuggestions.test.ts` |
| Regression | Existing stale-context locator tests still pass as fallback safety net | `ApplySuggestionCommandSearch.test.ts` |

## Migration / Rollout

No backend migration required. Implement in slices: (1) contracts + orchestrator tests, (2) mutation patch return path, (3) snapshot/rebase logic, (4) fallback/hot-paragraph tuning.

## Open Questions

- [ ] Should paragraph identity be stored by text span only, or by an additional stable local token in the snapshot?
- [ ] How much of the localized reread should happen inside `ApplySuggestionCommand` versus a dedicated snapshot-updater collaborator?
