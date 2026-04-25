# Proposal: Operational Wrapper Replace Resolution

## Intent
Productionize Word lab findings for replace resolution by establishing operational wrappers as a strong identity, enforcing `ambiguous-location` aborts, handling contiguous suggestions, and applying a comments-only cleanup policy. Legacy accept/reject workflow preservation is NOT a design goal. We treat legacy code as technical debt; if the current workflow is conceptually wrong, replacing/removing it is preferred over incremental patching.

## Scope

### In Scope
- Implement operational wrappers as the explicit resolution identity.
- Enforce explicit `ambiguous-location` aborts instead of legacy fallback heuristics.
- Process contiguous/adjacent replace suggestions as explicit groups (validating contiguous `acceptAll`/`rejectAll`).
- Apply a targeted comments-only cleanup policy.
- Ensure non-Stylistic tracked changes remain untouched.
- Clean architectural replacement; backward compatibility with the current broken flow should NOT be optimized for by default.
- Make absence of legacy accept/reject workflow code a mandatory, reviewable deliverable. This SDD is not complete if legacy paths remain hidden behind new branches, compatibility shims, dead code, or fallback helpers.

### Out of Scope
- Mixed decisions inside contiguous groups (these will intentionally degrade to `unobservable`).
- Broad production guarantees for stable word gaps (awaiting more evidence).
- Incremental patching of the existing `compound-v2` observation workflow.
- Backward compatibility with legacy suggestion states if they violate the new operational model.
- Keeping legacy accept/reject code “temporarily” for safety. If a legacy path contradicts the operational-wrapper model, it must be deleted or made explicitly unreachable by removing its call sites and tests.

## Approach
We will architect a clean replacement based on first-class operational lineage rather than cautiously extending the existing pipeline. Incremental patching is NOT the default strategy if a clean replacement is architecturally better. We will introduce explicit wrapper metadata and strict contiguous grouping contracts, aggressively tearing out the legacy ad-hoc observation and duplicate-CC scoring logic.

The implementation must be auditable for deletion, not merely behavior. Reviewers must be able to point at the final diff and verify that legacy accept/reject paths, fallback heuristics, and compatibility branches were removed rather than left dormant. If a class/function remains, it must have a current responsibility in the operational-wrapper workflow; otherwise it is legacy debt and fails the SDD.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `ApplySuggestionCommand.ts` | Modified | Inject strict wrapper identity; remove legacy full-body fallback. |
| `ReplaceIdentityParser.ts` | Replaced | Discard legacy scoring; enforce strict operational wrapper validation. |
| `SuggestionLocator.ts` | Replaced | Eliminate legacy candidate ranking in favor of contiguous cluster selection. |
| `SuggestionResolutionObserver.ts` | Modified | Remove legacy semantic-side grouping; implement strict contiguous cluster mapping. |

## Mandatory No-Legacy Validation

This change has a non-negotiable validation gate: after implementation, there must be no legacy accept/reject workflow code left in the active production path.

Validation must include all of the following:

- **Static code review:** identify and remove legacy fallback branches, duplicate-CC scoring, body-level rescue paths, and compatibility shims that preserve the old workflow.
- **Call-site review:** prove removed legacy helpers are no longer reachable from taskpane, workflow, adapter, command, observer, resolver, executor, or cleanup paths.
- **Test review:** delete or rewrite tests that encode legacy behavior; tests must assert the new operational-wrapper model, not old compatibility behavior.
- **Negative checks:** verify that unsupported old states degrade explicitly instead of routing through a preserved fallback.
- **Diff accountability:** final implementation notes must list which legacy files/functions/branches were deleted or replaced.

If any old workflow branch remains because it is still needed, the implementation must explicitly justify why it is not legacy. “Kept just in case” is not an acceptable justification.

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Breaking in-flight legacy suggestions | High | Accept the break. Legacy code is technical debt; we do not preserve broken flows. |
| Unrelated tracked changes grouped falsely | Medium | Enforce strict boundaries based solely on explicit wrapper identity. |
| Aggressive cleanup hides host state | Medium | Restrict cleanup explicitly to replace-adjacent comment residue. |
| Legacy code survives as dormant fallback | High | Treat as SDD failure. Verification must inspect code structure, call sites, and tests, not only runtime behavior. |

## Migration / Rollout
Direct cutover. We will not maintain dual-read migration paths. Backward compatibility with the current broken flow should NOT be optimized for by default. Legacy fallback branches will be removed entirely.

Rollout is considered successful only when old accept/reject workflow code is removed from the active codebase. A feature flag that keeps legacy behavior available, a hidden compatibility path, or a “temporary” fallback branch is not a valid rollout strategy for this change.

## Success Criteria

- [ ] **Slice 1:** Replaced legacy identity parsing with explicit operational wrappers and `ambiguous-location` aborts.
- [ ] **Slice 2:** Implemented contiguous group observation without legacy fallback logic.
- [ ] **Slice 3:** Added comments-only cleanup policy, leaving non-Stylistic changes untouched.
- [ ] **No-legacy verification:** Legacy workflows are fully removed, not just bypassed.
- [ ] **No dormant fallback verification:** No preserved legacy accept/reject branch remains reachable or hidden behind compatibility flags/shims.
- [ ] **Test-suite verification:** Tests that asserted legacy behavior are deleted or rewritten to assert operational-wrapper behavior.
- [ ] **Implementation report:** Final implementation identifies the removed/replaced legacy code paths and explains any remaining similarly named code by current responsibility.
