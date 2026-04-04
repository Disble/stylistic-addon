# stylistic-addon — Claude Instructions

## Skills (Auto-load based on context)

When working in this project, IMMEDIATELY read the corresponding skill file before writing any code or tests.

| Context | Read this file |
| ------- | -------------- |
| Any code in this project (adapters, domain, infrastructure) | `~/.claude/skills/stylistic-addon-architecture/SKILL.md` |
| Any test file (`.test.ts`), Office.js mocks, Word API assertions | `.claude/skills/stylistic-addon-testing/SKILL.md` |
| Debugging a regression, bug investigation, Word behavior mismatch | `.claude/skills/stylistic-addon-debugging/SKILL.md` |

Update the local testing skill when a regression teaches a new repo-specific lesson.
