## Exploration: Operational Wrapper Replace Resolution

### Current State
Current production replace resolution already runs through a dedicated workflow: `ApplySuggestionCommand` writes `compound-v2` metadata for replace suggestions, `SuggestionLocator` selects valid `compound-v2` content controls by tag, `SuggestionResolutionObserver` gathers evidence from content-control, range, body-related, operational-anchor, and comment surfaces, and `ReplaceResolutionWorkflow` executes semantic sides in action-dependent order.

Important current invariants are already in place: a replace suggestion is not treated as a single inserted-side wrapper, `0 tracked changes observed` degrades to `unobservable`, corrupt `compound-v2` degrades to `identity-lost`, and `SuggestionResolutionWorkflow` skips feedback unless the result is terminal `accepted` or `rejected`.

Architecturally, the gaps relevant to the lab themes are:
- wrapper identity is only implicit in `compound-v2` plus duplicate-CC scoring, not modeled as first-class operational lineage
- ambiguous-location aborts exist on the apply side (`ApplySuggestionCommand` returns `null` on ambiguous full-body fallback), but that policy is not yet elevated into an explicit replace-resolution production contract
- tracked changes are grouped only by semantic side (`Added` / `Deleted`), not by contiguous cluster or wrapper neighborhood
- comment cleanup is split between per-suggestion resolution cleanup and bulk `cleanupResolvedComments()`, with no explicit production rule for comments-only residue discovered during replace workflows

No `openspec/config.yaml` exists in the repo today, so this exploration follows the existing repository convention used by other active changes plus the shared OpenSpec `exploration.md` convention.

### Affected Areas
- `src/adapters/word/ApplySuggestionCommand.ts` — writes `compound-v2`, removes existing wrappers, and already aborts ambiguous full-body anchor fallback.
- `src/adapters/word/ReplaceIdentityParser.ts` — validates/scoring logic for persisted replace identity and the natural extension point for wrapper-related identity rules.
- `src/adapters/word/resolve-suggestion/SuggestionLocator.ts` — ranks duplicate CC candidates and is the right place for wrapper-level selection policy.
- `src/adapters/word/resolve-suggestion/SuggestionResolutionObserver.ts` — collects evidence sources and currently groups candidates only by semantic side.
- `src/adapters/word/resolve-suggestion/ResolutionContext.ts` — current observation contract would need expansion if contiguous grouping becomes first-class.
- `src/adapters/word/resolve-suggestion/ResolveSuggestionTrackChangeOrchestrator.ts` — consumes semantic candidates and would need to honor any stronger grouping/abort rules.
- `src/adapters/word/resolve-suggestion/SuggestionResolutionCleanup.ts` — owns per-suggestion cleanup after certified resolution.
- `src/adapters/word/cleanup/CommentCleanup.ts` — owns bulk resolved-comment cleanup and is the natural boundary for comments-only cleanup policy.
- `src/domain/suggestion/SuggestionResolutionWorkflow.ts` — already owns feedback suppression for ambiguous states and should stay the app-level owner of terminal semantics.
- `openspec/specs/replace-resolution-workflow/spec.md` and `openspec/changes/replace-suggestion-identity/*` — define the current replace identity/resolution baseline this new change should relate to without reusing.

### Approaches
1. **Extend the current `compound-v2` workflow** — keep the existing replace pipeline and add wrapper-aware selection, contiguous candidate grouping, and stricter cleanup/abort rules inside current adapter collaborators.
   - Pros: builds on production code already aligned with `unobservable` / `identity-lost`; lowest migration cost; keeps taskpane/domain churn small.
   - Cons: may stretch `compound-v2` metadata and current observer contracts; contiguous grouping must be added carefully to avoid bloating `SuggestionResolutionObserver`.
   - Effort: Medium

2. **Introduce a new identity/versioned workflow (`compound-v3` style)** — add explicit wrapper lineage and grouping metadata, then teach apply/locate/observe/cleanup to prefer the new model.
   - Pros: cleaner long-term model for wrapper identity and operational grouping; easier to express lab findings explicitly.
   - Cons: dual-read migration complexity; higher regression risk in apply, navigation, and resolution; likely overkill before proving current metadata is insufficient.
   - Effort: High

### Recommendation
Use Approach 1. This should be a follow-on productionization change that extends the current `compound-v2`-based replace workflow rather than replacing it. Concretely: keep domain/taskpane semantics as they are, add wrapper-aware policy in `ApplySuggestionCommand` + `ReplaceIdentityParser` + `SuggestionLocator`, introduce contiguous-group normalization as a dedicated resolution collaborator instead of more ad-hoc logic in `SuggestionResolutionObserver`, and define a conservative comments-only cleanup rule inside the existing cleanup boundary.

This change should relate to `replace-suggestion-identity` as a downstream hardening step, not as a continuation of that change. The existing change established the identity/status contract; this new change should operationalize real-Word lab findings around wrapper selection, ambiguity aborts, cluster selection, and cleanup safety.

Likely implementation slices after proposal/spec:
1. Codify wrapper-identity and ambiguity-abort rules.
2. Add contiguous-group observation/selection contract for replace evidence.
3. Add comments-only cleanup policy for replace-adjacent residue.
4. Add focused adapter/taskpane regressions only where public behavior changes.

### Risks
- Word may expose adjacent unrelated tracked changes that look contiguous, causing false grouping unless grouping stays suggestion-scoped.
- Wrapper-aware selection may still face indistinguishable duplicate CCs when metadata is structurally valid but operationally stale.
- Expanding observer contracts without extracting collaborators could worsen complexity in `SuggestionResolutionObserver.ts`.
- Cleanup policy is dangerous: deleting comments or wrappers too aggressively can hide unresolved host state.
- If lab findings actually require new persisted metadata, an incremental `compound-v2` extension may stall and force a later version bump anyway.

### Ready for Proposal
Yes — the next artifact should be `proposal.md` for `operational-wrapper-replace-resolution`, scoped as a production-hardening follow-up to the current replace identity workflow rather than a reuse of `replace-suggestion-identity`.
