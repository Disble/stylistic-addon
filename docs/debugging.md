# Debugging

This document is the canonical debugging workflow for `stylistic-addon`.

Complementary documents:

- [`troubleshooting.md`](./troubleshooting.md) — current symptom → cause → fix
  guidance.
- [`replace-resolution-postmortem.md`](./replace-resolution-postmortem.md) —
  replace-resolution forensic history.
- [`architecture.md`](./architecture.md) — current module boundaries and
  collaborator contracts.
- [`testing.md`](./testing.md) — how to choose the right regression tests while
  debugging.

## Evidence-first workflow

Treat bug reports as first-class evidence, especially when they come with:

- screenshots,
- explicit current behavior,
- explicit expected behavior,
- real-host logs.

Do not rush into implementation before you understand what the evidence proves.

## Investigation ladder

Use this order unless the evidence clearly skips a layer:

1. Observed symptom.
2. Taskpane/UI state.
3. Adapter/workflow result contract.
4. Word creation/resolution mechanics.
5. Office.js host semantics.

This prevents getting trapped in one file too early.

## Debugging rules

- Real Word beats the first GREEN.
- Fix the test model if the host disproves it.
- Split tests surgically on the active bug path; do not refactor the whole suite
  unless the task is explicitly test cleanup.
- When a bug reveals that a previous fix or assumption was incomplete, update the
  relevant docs so the correction survives the session.
