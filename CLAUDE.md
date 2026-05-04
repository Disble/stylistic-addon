# stylistic-addon — Claude Instructions

## Skills (Auto-load based on context)

When working in this project, IMMEDIATELY read the corresponding skill file before writing any code or tests.

| Context | Read this file |
| ------- | -------------- |
| Any code in this project (adapters, domain, infrastructure) | `~/.claude/skills/stylistic-addon-architecture/SKILL.md` |
| Any test file (`.test.ts`), Office.js mocks, Word API assertions | `.claude/skills/stylistic-addon-testing/SKILL.md` |
| Debugging a regression, bug investigation, Word behavior mismatch | `.claude/skills/stylistic-addon-debugging/SKILL.md` |

Update the local testing skill when a regression teaches a new repo-specific lesson.

- `docs/` is the canonical home for detailed repo-specific bug forensics and operational lessons.
- When a lesson would duplicate project docs, update the relevant `docs/` page first and keep the skill as a concise pointer/checklist.

## Philosophy

- This project ALWAYS prioritizes breaking changes over legacy code.
- This project consider legacy code as TECHNICAL DEBT to be paid off, not a constraint to work around.
- Each new feauture, bug fix, or refactor should be designed from the start with SOLID design patterns, clean architecture principles, and best testing practices in mind.
