# Skill Registry

Generated for project: `stylistic-addon`

## Available Skills

| Name | Scope | Path | Trigger / Use When |
|---|---|---|---|
| `mastra` | project | `.claude/skills/mastra/SKILL.md` | Use for any Mastra development; verify current APIs/docs before touching `@mastra/*`. |
| `kin` | user | `C:/Users/User/.config/opencode/skills/kin/SKILL.md` | Use when code touches external libraries, library stack traces appear, or API signatures are uncertain. |
| `kin-init` | user | `C:/Users/User/.config/opencode/skills/kin-init/SKILL.md` | Use to register or repair KIN integration and docs routing. |
| `cognitive-complexity` | user | `C:/Users/User/.config/opencode/skills/cognitive-complexity/SKILL.md` | Use when measuring or reducing Cognitive Complexity. |
| `go-testing` | user | `C:/Users/User/.config/opencode/skills/go-testing/SKILL.md` | Use when writing Go tests, especially Bubbletea/TUI tests. |
| `react-doctor` | user | `C:/Users/User/.config/opencode/skills/react-doctor/SKILL.md` | Run after React changes to catch correctness, performance, and architecture issues. |
| `skill-creator` | user | `C:/Users/User/.config/opencode/skills/skill-creator/SKILL.md` | Use when creating new AI skills or documenting reusable agent workflows. |

## Project Instruction Files

No supported project instruction index files were found at the root (`AGENTS.md`, `agents.md`, `CLAUDE.md`, `.cursorrules`, `GEMINI.md`, `copilot-instructions.md`).

## Notes

- SDD skills were intentionally excluded from this registry per bootstrap rules.
- Project-level skills override user-level skills with the same name.
- KIN routing is configured in `.atl/docs-routing.yaml` and should be read before resolving external-library documentation.
