---
name: stylistic-addon-architecture
description: >
  Router skill for the stylistic-addon architecture docs.
  Trigger: When writing code in this project, adding features, creating adapters,
  ports, workflows, or modifying taskpane composition.
license: Apache-2.0
metadata:
  author: gentleman-programming
  version: "2.0"
---

# Skill: stylistic-addon-architecture

This skill is a **router**, not the canonical architecture spec.

Before writing code in `stylistic-addon`, read the project docs that match the
change you are making.

## Canonical docs

- `docs/architecture.md` — current system design, boundaries, collaborators, and
  composition rules.
- `docs/review-domain-and-track-changes.md` — frontend domain definition,
  ubiquitous language, and Track Changes lifecycle policy.
- `docs/linting-and-file-anatomy.md` — enforced file anatomy, component folder
  contracts, `__tests__/` placement, and runtime-module rules.
- `docs/api-contract.md` — backend/Mastra contract, workflow shape, and auth
  expectations.
- `docs/troubleshooting.md` — host quirks, auth/debugging gotchas, and known
  operational pitfalls.

## Mandatory reading by change type

### Any production code change

Read:

1. `docs/architecture.md`
2. `docs/linting-and-file-anatomy.md`

### Word review / accept-reject / Track Changes work

Also read:

- `docs/review-domain-and-track-changes.md`
- `docs/replace-resolution-postmortem.md` when the change touches replace
  resolution or Word observation semantics.

### Backend workflow / auth integration work

Also read:

- `docs/api-contract.md`
- `docs/troubleshooting.md`

## Non-negotiables

- React is allowed only under `src/taskpane/**`.
- `src/taskpane/taskpane.ts` is the Office composition root.
- `src/taskpane/index.tsx` is React + Fluent bootstrap only.
- Follow the enforced anatomy in `docs/linting-and-file-anatomy.md`; do not
  invent alternate file shapes.
- Use Bun commands (`bun run ...`) in docs/examples for this repo.

If a rule here and a rule in `/docs` disagree, `/docs` wins and this skill must
be updated.
