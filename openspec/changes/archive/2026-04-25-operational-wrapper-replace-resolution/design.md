# Design: Operational Wrapper Replace Resolution

## Technical Approach

Replace the current replace accept/reject skeleton with a strict operational-wrapper workflow. `ResolveSuggestionCommand` stays the entrypoint, but replace resolution proceeds only when a persisted wrapper can be located uniquely, expanded into one contiguous Stylistic group when needed, and observed safely before any accept/reject mutation. Comments and inserted-side content controls are metadata only; wrapper lineage is the semantic identity. No-legacy is architectural, not aspirational: duplicate-CC scoring, body-level rescue heuristics, and replace CC deletion cleanup must be removed, not hidden.

## Goals / Non-Goals

- MUST enforce `ambiguous-location` abort before mutation.
- MUST resolve contiguous/adjacent replace suggestions as one all-or-nothing group.
- MUST degrade mixed decisions inside a contiguous group to `unobservable`.
- MUST keep cleanup comments-only and preserve non-Stylistic tracked changes.
- MUST NOT preserve `compound-v2` ranking, full-body fallback, or dormant compatibility branches.
- MUST NOT generalize the lab one-word-gap result into a universal production heuristic.

## Architecture Decisions

| Decision | Choice | Rationale |
|---|---|---|
| Semantic identity | Persist an operational wrapper contract with group metadata | Replace resolution depends on lineage across both semantic sides, not one inserted-side CC |
| Grouping | Introduce `OperationalWrapperGroupResolver` for contiguous clusters only | Group execution must be explicit and auditable, not score-based |
| Abort policy | Raise `ambiguous-location` before `TrackedChangeResolutionExecutor.apply()` or `insertText(...replace)` | Fail-closed avoids silent corruption |
| Cleanup | Delete only Stylistic comments after replace resolution | `ContentControl.delete(false/true)` on replace metadata is lab-proven unsafe |
| Rollout | Direct cutover with explicit degradation of unsupported states | Any legacy fallback path fails the change goal |

## Component Design

```text
ResolveSuggestionCommand
  -> SuggestionLocator
  -> OperationalWrapperGroupResolver
  -> SuggestionResolutionObserver
  -> ReplaceResolutionWorkflow
  -> TrackedChangeResolutionExecutor
  -> SuggestionResolutionCleanup / CommentCleanup
```

```ts
type ResolutionAbortReason = "ambiguous-location" | "mixed-group";
interface OperationalReplaceMember {
  wrapperTag: string; groupId: string; groupIndex: number;
}
```

`ApplySuggestionCommand` MUST write wrapper/group identity for new replace suggestions and MUST stop when location would require ambiguous tie-breaking. `SuggestionResolutionObserver` MUST return only Stylistic tracked changes tied to the resolved wrapper/group.

## Lab Implementation Reference (read-only)

The lab repo `D:\dev\eln\stylistic\reject-word-testing` is a read-only evidence source, NOT a production dependency. Implementers should copy the semantic pattern only; they MUST NOT copy the lab taskpane UI, debug logging, scenario buttons, or manual harness scaffolding into production. The lab exists to help rebuild correctly and DOES NOT justify keeping legacy production code alive.

Reference points:

- `BITACORA_REJECT_TRACKED_CHANGES.md` — Scenario 19-23 evidence and final production rule.
- `src/taskpane/taskpane.ts` — concrete mechanics: `createOperationalWrapperInContext()`, `createGroupedOperationalWrapperInContext()`, `createAddonLikeReplaceInsideExistingWrapper()`, `deleteArtifactComments()`, plus scenario runners `runAmbiguousContextAbortScenario()`, `runAdjacentGroupedWrapperScenario()`, `runSafeCleanupCommentsOnlyPolicyScenario()`, `runOneWordGapIndependentScenario()`, and `runGroupedAcceptAllScenario()`.
- `src/taskpane/taskpane.html` — scenario labels only; UI wiring is lab-only scaffolding.

Scenario-to-design mapping:

- **19** → abort with `ambiguous-location` before wrapper creation or replace mutation.
- **20** → contiguous adjacent suggestions use one grouped wrapper and support grouped `rejectAll()`.
- **21** → cleanup is comments-only; wrapper and inserted-side CC persist until a separately validated cleanup exists.
- **22** → one-word-gap independent wrappers passed in the lab, but this is evidence of a narrow safe case, NOT a universal distance heuristic.
- **23** → the same grouped-wrapper model supports grouped `acceptAll()`.

Unvalidated area remains explicit: mixed decisions inside a contiguous group are still degraded/unobservable; do not overclaim beyond grouped all-accept or all-reject.

## File Changes

| File | Action | Description |
|---|---|---|
| `src/domain/types.ts` | Modify | Add wrapper/group metadata and `ResolutionAbortReason` |
| `src/adapters/word/ApplySuggestionCommand.ts` | Modify | Persist wrapper identity; remove replace full-body rescue/scoring |
| `src/adapters/word/ReplaceIdentityParser.ts` | Replace | Strict parser/validator only; delete scoring helpers |
| `src/adapters/word/resolution/SuggestionLocator.ts` | Replace | Exact wrapper lookup with ambiguity detection, no ranking |
| `src/adapters/word/resolution/OperationalWrapperGroupResolver.ts` | Create | Expand contiguous group from the seed wrapper |
| `src/adapters/word/resolution/SuggestionResolutionObserver.ts` | Modify | Observe groups, mixed decisions, and non-Stylistic preservation |
| `src/adapters/word/resolution/ReplaceResolutionOrchestrator.ts` | Modify | Execute all-or-nothing group strategy |
| `src/adapters/word/resolution/SuggestionResolutionCleanup.ts` / `src/adapters/word/cleanup/CommentCleanup.ts` | Modify | Enforce comments-only cleanup; make replace CC deletion unreachable |

## Verification / No-Legacy Validation

- Remove `scoreCompoundReplaceIdentityMatch`, `rankResolutionContentControls`, replace ranking call sites, and replace-specific body rescue branches.
- Prove no active path from command, locator, observer, orchestrator, executor, or cleanup reaches legacy fallback logic.
- Rewrite tests around exact wrapper lookup, grouped resolution, mixed-group degradation, comments-only cleanup, and non-Stylistic preservation; delete tests that bless ranking/fallback behavior.
- Fail the change if any reachable legacy fallback, replace CC deletion path, or preserved compatibility branch remains.

## Migration / Risks

Direct cutover only. Old artifacts that do not satisfy the new wrapper contract degrade to `identity-lost` or `unobservable`; they are not migrated through compatibility code. Accepted risks: in-flight legacy suggestions break; comments-only cleanup leaves metadata longer; grouping stays limited to explicit wrapper metadata plus Word adjacency/overlap, not raw gap counting.
