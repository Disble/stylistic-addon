/**
 * Claude Code PostToolUse hook — Write|Edit
 *
 * After any file edit, runs the co-located Vitest test suite if one exists.
 * Skips silently for non-.ts files or files with no co-located .test.ts.
 * Exits with code 1 if tests fail, blocking Claude Code from continuing.
 */

import { execSync } from "node:child_process";
import { existsSync } from "node:fs";

const PROJECT_ROOT = "D:/dev/eln/stylistic/stylistic-addon";

let raw = "";
process.stdin.on("data", (chunk) => (raw += chunk));
process.stdin.on("end", () => {
  const payload = JSON.parse(raw);
  const filePath = payload?.tool_input?.file_path ?? "";

  // Skip non-TypeScript files
  if (!filePath?.endsWith(".ts")) process.exit(0);

  // Resolve the test file: if already a .test.ts, run it directly;
  // otherwise look for a co-located .test.ts
  const testFile = filePath.endsWith(".test.ts")
    ? filePath
    : filePath.replace(/\.ts$/, ".test.ts");

  if (!existsSync(testFile)) process.exit(0);

  try {
    execSync(`npx vitest run "${testFile}"`, {
      cwd: PROJECT_ROOT,
      stdio: "inherit",
    });
  } catch {
    process.exit(1);
  }
});
