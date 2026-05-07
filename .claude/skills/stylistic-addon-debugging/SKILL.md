---
name: stylistic-addon-debugging
description: >
  Router skill for stylistic-addon debugging docs.
  Trigger: When investigating regressions, Word/Office.js behavior mismatches,
  or bugs that escaped tests.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "2.0"
---

# Skill: stylistic-addon-debugging

This skill is a **router** to the canonical debugging docs.

## Read first

- `docs/debugging.md` — debugging workflow, intake ladder, and evidence-first
  process.
- `docs/troubleshooting.md` — current symptom → cause → inspection guidance.
- `docs/replace-resolution-postmortem.md` — replace-resolution forensic history.
- `docs/architecture.md` — current collaborator boundaries before changing code.

## Use this skill when

- real Word behavior contradicts tests,
- a regression escaped existing coverage,
- you need to decide whether the bug is UI, workflow, adapter, or host-level.

## Operating rule

Do not duplicate repo forensics here. Update `/docs` first, then keep this skill
as a thin router/checklist.
