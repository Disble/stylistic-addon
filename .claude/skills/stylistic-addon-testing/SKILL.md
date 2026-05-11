---
name: stylistic-addon-testing
description: >
  Router skill for stylistic-addon testing docs.
  Trigger: When writing, reviewing, or refactoring tests, especially Office.js
  mocks, taskpane tests, and Word adapter regression coverage.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "2.0"
---

# Skill: stylistic-addon-testing

This skill is a **router**, not the canonical testing spec.

## Canonical docs

- `docs/testing.md` — testing strategy, confidence tiers, mock discipline, and
  verification commands.
- `docs/linting-and-file-anatomy.md` — enforced `__tests__/` placement and test
  helper anatomy.
- `docs/troubleshooting.md` — real-host discrepancies and operational gotchas.
- `docs/replace-resolution-postmortem.md` — replace-resolution regression lessons.
- `docs/architecture.md` — current workflow and adapter contracts under test.

## Mandatory defaults

- Prefer behavior/contract tests over choreography tests.
- Real Word beats fake GREEN.
- Keep strict Office.js mocks when the bug lives at the host boundary.
- Put tests under sibling `__tests__/` directories.
- Use Bun commands for examples and verification.

If testing guidance evolves, update `/docs` first and keep this skill as a short
entrypoint.
