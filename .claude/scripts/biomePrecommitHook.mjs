/**
 * Claude Code PreToolUse hook — Bash (git commit)
 *
 * Before any git commit, runs `biome check --error-on-warnings` on staged
 * TypeScript/CSS/JSON files. If warnings are found, blocks the commit with
 * a permissionDecision: "deny" response so Claude Code cannot proceed.
 *
 * This ensures the agent NEVER commits code with Biome warnings.
 */

import { execSync } from "node:child_process";

const PROJECT_ROOT = "D:/dev/eln/stylistic/stylistic-addon";

let raw = "";
process.stdin.on("data", (chunk) => (raw += chunk));
process.stdin.on("end", () => {
  const payload = JSON.parse(raw);
  const command = payload?.tool_input?.command ?? "";

  // Only intercept git commit commands
  if (!command.includes("git commit")) process.exit(0);

  try {
    // Get list of staged files
    const stagedOutput = execSync("git diff --staged --name-only", {
      cwd: PROJECT_ROOT,
    })
      .toString()
      .trim();

    if (!stagedOutput) process.exit(0);

    const stagedFiles = stagedOutput
      .split("\n")
      .filter((f) => /\.(ts|css|json)$/.test(f));

    if (stagedFiles.length === 0) process.exit(0);

    // Run biome with --error-on-warnings — exits non-zero if ANY warning exists
    execSync(`npx biome check --error-on-warnings ${stagedFiles.join(" ")}`, {
      cwd: PROJECT_ROOT,
      stdio: "pipe",
    });

    // All clean — allow the commit
    process.exit(0);
  } catch {
    // Warnings found — deny the commit
    process.stdout.write(
      JSON.stringify({
        hookSpecificOutput: {
          hookEventName: "PreToolUse",
          permissionDecision: "deny",
          permissionDecisionReason:
            "Biome warnings detected in staged files. Fix ALL warnings before committing. Run: npx biome check --write --unsafe",
        },
      }),
    );
    process.exit(1);
  }
});
