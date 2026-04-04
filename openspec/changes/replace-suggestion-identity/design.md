# Design: Replace Suggestion Identity

## Technical Approach

Current model: `ApplySuggestionCommand` creates one inserted-side `ContentControl`; `WordAdapter` later infers the whole replace via `cc.getTrackedChanges()` + `body.getTrackedChanges()` + spatial overlap. Proposed model: one `ReviewSuggestion` keeps one domain identity and multiple `WordArtifactRef`s. Business status is computed only after an explicit observation step.

Phase A is defensive: `0 tracked changes observed` MUST stop mapping to `already-resolved`. Phase B introduces compound replace identity for newly applied suggestions and removes `legacy-v1` as a supported replace-resolution path.

## Architecture Decisions

| Decision | Choice | Alternatives | Rationale |
|---|---|---|---|
| Observation vs business status | Split them explicitly | Keep one flat status enum | Prevents epistemic bug: host invisibility is not resolution |
| New states | Add `unobservable` and `identity-lost`; keep `already-resolved` only for positive confirmation | Reuse `error` or `already-resolved` | `unobservable` = host could not prove state; `identity-lost` = v2 identity is incomplete/corrupt |
| Replace identity | Compound identity with inserted-side ref, deleted-side ref, anchor ref | Keep CC as full identity | Replace is semantically two-sided in Word |
| Rollout | Two phases with hard cut to `compound-v2` in Phase B | Long-lived dual-read legacy/v2 | Stops false positives fast without keeping the broken model alive |

## Data Flow

```text
taskpane → ReviewSessionMediator → SuggestionResolutionWorkflow → IDocumentPort
                                                           ↓
                                  observe replace artifacts → conclude status

Current: CC(ref) → infer replace → maybe "already-resolved"
Proposed: identity + refs → observe refs → observation status → business status
```

## Invariants

- A replace suggestion MUST NOT treat one inserted-side `ContentControl` as the whole identity.
- `already-resolved` MUST require positive evidence of prior accept/reject resolution.
- `0 tracked changes observed` MUST yield `unobservable`, never terminal success.
- V2 replace identity is fully observable only when inserted-side and deleted-side semantics are both accounted for.
- Feedback MUST be skipped for `unobservable` and `identity-lost`.

## Phased Design

### Phase A — Defensive immediate

- Modify `WordAdapter.resolveSuggestion()` so empty tracked-change observation returns `unobservable`.
- Keep CC/comment cleanup conservative: do not delete the anchor on `unobservable`.
- Extend workflow, suggestion state machine, mediator, and taskpane to render ambiguous state as retryable/non-terminal.

### Phase B — Compound identity

- Persist versioned replace identity metadata for newly applied replace suggestions.
- Separate in `WordAdapter`: locate refs → observe inserted/deleted sides → classify observation → mutate only when observable enough.
- Reserve `identity-lost` for v2 suggestions whose mandatory refs/metadata are missing or contradictory.

## Interfaces / Contracts

```ts
type SuggestionObservationStatus =
  | "confirmed-pending"
  | "confirmed-resolved"
  | "unobservable"
  | "identity-lost";

interface WordArtifactRef {
  kind: "content-control" | "tracked-change" | "comment" | "anchor";
  role: "inserted-side" | "deleted-side" | "operational-anchor";
  value: string;
}

interface ReplaceSuggestionIdentity {
  suggestionId: string;
  version: "compound-v2";
  insertedSideRef: WordArtifactRef;
  deletedSideRef?: WordArtifactRef;
  anchorRef?: WordArtifactRef;
}
```

`SuggestionActionResult.status` should add `unobservable | identity-lost`. `already-resolved` stays, but only for confirmed resolution.

## File Changes

| File | Action | Description |
|---|---|---|
| `src/domain/types.ts` | Modify | Add observation/identity types and new action statuses |
| `src/domain/ports.ts` | Modify | Preserve port shape but document stronger result contract |
| `src/adapters/word/ApplySuggestionCommand.ts` | Modify | Emit versioned replace metadata; stop implying CC = full identity |
| `src/adapters/word/WordAdapter.ts` | Modify | Split observation from resolution; require compound-v2 for replace resolution |
| `src/domain/suggestion/SuggestionResolutionWorkflow.ts` | Modify | Skip feedback for ambiguous observation states |
| `src/domain/suggestion/SuggestionStateMachine.ts` | Modify | Treat `unobservable` as non-terminal retryable; `identity-lost` likely terminal-warning |
| `src/domain/review/ReviewSessionMediator.ts` | Modify | Pass through richer resolution semantics |
| `src/taskpane/taskpane.ts` | Modify | Render ambiguous/identity-loss states explicitly |
| `src/adapters/word/*.test.ts`, `src/domain/**/*.test.ts`, `src/taskpane/*.test.ts` | Modify | Add regression coverage for phase A/B |

## Testing Strategy

| Layer | What to Test | Approach |
|---|---|---|
| Unit | `0 observed` → `unobservable` | `WordAdapterAccept/RejectSuggestion.test.ts` |
| Unit | feedback skipped on ambiguous states | `SuggestionResolutionWorkflow.test.ts` |
| Unit | valid state transitions for new states | `SuggestionStateMachine.test.ts` |
| Integration-ish | taskpane rendering for `unobservable` / `identity-lost` | `TaskpaneSuggestionResolution.test.ts` |
| Regression | reject bare-ID as unsupported and certify `compound-v2`/corrupt metadata paths | action test helpers + adapter tests |

## Migration / Rollout

Phase B adopts a hard cut for replace resolution:
- replace suggestions are written and resolved only through `compound-v2` metadata,
- bare-ID / `legacy-v1` replace artifacts are not treated as supported identities,
- corrupt or incomplete compound metadata downgrades to `identity-lost`,
- lack of observation downgrades to `unobservable`, never `already-resolved`.

## Risks / Tradeoffs / Open Questions

- Tradeoff: more states increase UI/test complexity, but they remove false terminal lies.
- Risk: Word may not expose stable tracked-change IDs across sessions; phase B should not depend on IDs alone.
- Open: exact v2 metadata serialization format in CC tag/title/comment payload.
- Open: whether `identity-lost` should allow retry or force re-analysis/navigation only.

## Recommended Implementation Sequence

1. Phase A statuses + tests in domain/adapter/workflow/taskpane.
2. Ship conservative behavior with feedback suppression.
3. Add v2 replace identity write path in `ApplySuggestionCommand`.
4. Add compound-v2 observation path in `WordAdapter`.
5. Add `identity-lost` handling and hard-cut regression matrix.
