# Apply Progress: Operational Wrapper Replace Resolution

> Historical apply artifact. Command strings below have been normalized to the
> current Bun-based workflow so this record does not mislead future readers.

## 2026-04-25 verification-fix pass

- Implemented real Word adjacency/contiguity validation for operational wrapper groups using `Range.compareLocationWith()` on consecutive group members.
- Added production `mixed-group` degradation when a contiguous group exposes incompatible per-member semantic sides before any executor mutation.
- Added grouped all-or-nothing accept/reject tests that assert all member tracked changes are resolved semantically.
- Added command/workflow/taskpane guardrails for `ambiguous-location` / `mixed-group` no-mutation and feedback-skipped behavior.
- Strengthened comments-only cleanup evidence with a shared document-graph harness: the selected Stylistic comment is removed from connected fixture state while wrapper CCs, inserted-side CCs, and foreign tracked changes remain untouched in the mock model. This is honest mock-level evidence, not a claim of real-Word exhaustiveness.
- Ran focused non-build verification: `bun run typecheck` and focused Vitest suites. All passed.
