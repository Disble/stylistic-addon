# Skill Registry

Generated for project: `stylistic-addon`

## Available Skills

| Name | Scope | Path | Trigger / Use When |
|---|---|---|---|
| `stylistic-addon-testing` | project | `.claude/skills/stylistic-addon-testing/SKILL.md` | Use for addon tests, Office.js mocks, tracked-change regressions, and confidence audits. |
| `mastra` | project | `.claude/skills/mastra/SKILL.md` | Use for any Mastra development; verify current APIs/docs before touching `@mastra/*`. |
| `kin` | user | `C:/Users/User/.config/opencode/skills/kin/SKILL.md` | Use when code touches external libraries, library stack traces appear, or API signatures are uncertain. |
| `kin-init` | user | `C:/Users/User/.config/opencode/skills/kin-init/SKILL.md` | Use to register or repair KIN integration and docs routing. |
| `cognitive-complexity` | user | `C:/Users/User/.config/opencode/skills/cognitive-complexity/SKILL.md` | Use when measuring or reducing Cognitive Complexity. |
| `go-testing` | user | `C:/Users/User/.config/opencode/skills/go-testing/SKILL.md` | Use when writing Go tests, especially Bubbletea/TUI tests. |
| `react-doctor` | user | `C:/Users/User/.config/opencode/skills/react-doctor/SKILL.md` | Run after React changes to catch correctness, performance, and architecture issues. |
| `skill-creator` | user | `C:/Users/User/.config/opencode/skills/skill-creator/SKILL.md` | Use when creating new AI skills or documenting reusable agent workflows. |

## Project Instruction Files

| File | Purpose |
|---|---|
| `AGENTS.md` | Project routing and mandatory skill-loading rules |
| `CLAUDE.md` | Claude-facing project entrypoint and testing-skill reminder |

## Notes

- SDD skills were intentionally excluded from this registry per bootstrap rules.
- Project-level skills override user-level skills with the same name.
- KIN routing is configured in `.atl/docs-routing.yaml` and should be read before resolving external-library documentation.
