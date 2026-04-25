# Archive Report: Operational Wrapper Replace Resolution

**Change**: `operational-wrapper-replace-resolution`  
**Archived On**: 2026-04-25  
**Final Status**: Archived after passed verification

## Executive Summary

- Shipped operational-wrapper replace resolution as the new strict production path.
- Preserved the shared replace semantic-order policy while intentionally deleting legacy fallback, scoring, ranking, and body-rescue behavior instead of keeping compatibility shims.
- Final verification passed, including no-legacy validation and sufficient repo/mock-level cleanup evidence.

## Source-of-Truth Sync

| Domain | Action | Details |
|---|---|---|
| `replace-resolution-workflow` | Updated | Added 5 requirements, preserved 1 existing requirement, removed the legacy workflow-skeleton requirement from the main spec. |

Main spec updated:

- `openspec/specs/replace-resolution-workflow/spec.md`

## Archived Outcome

### Strict operational-wrapper path shipped

The archived change establishes operational wrapper metadata as the explicit identity for replace resolution, enforces `ambiguous-location` aborts before mutation, and resolves contiguous wrapper groups as all-or-nothing `acceptAll()` / `rejectAll()` operations.

### Anti-legacy outcome is explicit

This archive records a direct cutover, not a compatibility bridge. The change intentionally removed or made unreachable legacy behavior, including:

- duplicate-CC scoring and ranking paths,
- full-body fallback and body-rescue discovery paths,
- semantic-side-only replace grouping,
- replace-wrapper/content-control deletion cleanup,
- dormant compatibility branches that would preserve the old replace workflow.

Legacy fallback/scoring/ranking/body-rescue paths were intentionally removed rather than preserved for safety. That anti-legacy decision remains part of the archived source of truth.

## Verification Basis

- `verify-report.md` status: **passed**
- No-legacy validation: **passed**
- Cleanup validation: **passed**
- Tasks truthfulness: **passed**

Supporting evidence reviewed directly during archive:

- `exploration.md`
- `proposal.md`
- `design.md`
- `specs/replace-resolution-workflow/spec.md`
- `tasks.md`
- `apply-progress.md`
- `verify-report.md`

## Cleanup Confidence Notes

Cleanup evidence is sufficient at repo/mock level because the connected cleanup harness proves the real cleanup path removes only the targeted Stylistic comment while preserving wrapper content controls, inserted-side metadata, unrelated comments, and foreign tracked changes.

This is NOT a claim that mocks overrule Word host behavior. Real Microsoft Word remains the ultimate authority for host behavior, so this note is archived as a future-confidence note, not a release blocker.

## Archive Contents Checklist

- `exploration.md` ✅
- `proposal.md` ✅
- `design.md` ✅
- `specs/replace-resolution-workflow/spec.md` ✅
- `tasks.md` ✅
- `apply-progress.md` ✅
- `verify-report.md` ✅
- `archive-report.md` ✅

## Closure

The SDD cycle for `operational-wrapper-replace-resolution` is complete: explored, proposed, designed, implemented, verified, synced into the main spec, and archived.
