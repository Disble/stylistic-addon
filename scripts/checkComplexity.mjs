import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";

// ---------------------------------------------------------------------------
// Thresholds
// ---------------------------------------------------------------------------

const LINE_WARN = 400;
const LINE_ERROR = 500;

const TEST_LINE_WARN = 500;
const TEST_LINE_ERROR = 800;

/**
 * Known exceptions — files whose complexity is inherent, not a design smell.
 * Each entry acts as a RATCHET: the allowed max should DECREASE over time
 * as refactors chip away at the file. Raising a cap requires justification.
 */
const KNOWN_EXCEPTIONS = new Map([]);

/**
 * Known exceptions for test/helper files — same ratchet contract as above.
 */
const KNOWN_EXCEPTIONS_TEST = new Map([
  [
    "src/taskpane/TaskpaneTestHelper.ts",
    {
      maxLines: 650,
      reason:
        "Temporary exception while TaskpaneTestHelper is split into focused per-concern helpers",
    },
  ],
]);

/**
 * Per-folder refactoring guidance — shown on warn/error results.
 * Ordered most-specific first; first match wins.
 */
const FOLDER_GUIDANCE = [
  [
    "src/adapters/word/",
    "Word adapter boundary — docs/architecture.md §2.1 §3.3. " +
      "Extract into a focused collaborator using approved suffixes: Command · Orchestrator · Cleanup · Locator · Observer · Factory.",
  ],
  [
    "src/adapters/mastra/",
    "Backend adapter boundary — docs/architecture.md §2.1. " +
      "One class per transport concern; extract mapping, retry, or polling into separate collaborators.",
  ],
  [
    "src/adapters/",
    "Adapter boundary — docs/architecture.md §2.1 (ports and adapters rule). " +
      "Each adapter implements one port. Extract growing responsibilities into focused collaborators.",
  ],
  [
    "src/domain/",
    "Domain layer — docs/architecture.md §2.1. Zero framework dependencies. " +
      "Extract into handlers, state machines, or value objects. Never import adapters or Office.js.",
  ],
  [
    "src/taskpane/",
    "Presentation layer — docs/architecture.md §9.2. " +
      "Taskpane consumes ports, mediator, and workflows only. Extract rendering or event-handling logic.",
  ],
  [
    "src/infrastructure/",
    "Infrastructure layer — pure functions and constants only. " +
      "No application logic should accumulate here. Extract into focused pure utility modules.",
  ],
];

function getFolderGuidance(relativePath) {
  for (const [prefix, guidance] of FOLDER_GUIDANCE) {
    if (relativePath.startsWith(prefix)) return guidance;
  }
  return null;
}

/** File patterns exempt from line-count checks (structurally large). */
const EXEMPT_PATTERNS = [/types\.ts$/, /\.d\.ts$/];

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const rootDirectory = process.cwd();
const sourceDirectory = path.join(rootDirectory, "src");
const managedRoots = [
  path.join(sourceDirectory, "domain"),
  path.join(sourceDirectory, "adapters"),
  path.join(sourceDirectory, "infrastructure"),
  path.join(sourceDirectory, "taskpane"),
];
const supportedExtensions = [".ts", ".tsx"];

// ---------------------------------------------------------------------------
// File Collection
// ---------------------------------------------------------------------------

async function collectFiles(directory) {
  let entries;
  try {
    entries = await readdir(directory, { withFileTypes: true });
  } catch {
    return [];
  }

  const files = [];
  for (const entry of entries) {
    const entryPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await collectFiles(entryPath)));
      continue;
    }
    if (!entry.isFile()) continue;
    if (entry.name.endsWith(".d.ts")) continue;
    if (supportedExtensions.some((ext) => entry.name.endsWith(ext))) {
      files.push(entryPath);
    }
  }
  return files;
}

function isManagedFile(filePath) {
  return managedRoots.some((root) => filePath.startsWith(`${root}${path.sep}`));
}

function isTestFile(relativePath) {
  return (
    relativePath.endsWith(".test.ts") ||
    relativePath.endsWith(".spec.ts") ||
    relativePath.includes("TestHelper")
  );
}

// ---------------------------------------------------------------------------
// Analysis
// ---------------------------------------------------------------------------

async function analyzeFile(filePath) {
  const content = await readFile(filePath, "utf8");
  const relativePath = path
    .relative(rootDirectory, filePath)
    .replace(/\\/g, "/");
  const lineCount = content.split("\n").length;

  return { relativePath, lineCount };
}

function checkLineCount(file) {
  const { relativePath, lineCount } = file;

  if (EXEMPT_PATTERNS.some((pattern) => pattern.test(relativePath))) {
    return null;
  }

  const isTest = isTestFile(relativePath);
  const warnThreshold = isTest ? TEST_LINE_WARN : LINE_WARN;
  const errorThreshold = isTest ? TEST_LINE_ERROR : LINE_ERROR;
  const exceptions = isTest ? KNOWN_EXCEPTIONS_TEST : KNOWN_EXCEPTIONS;

  const exception = exceptions.get(relativePath);
  if (exception) {
    if (lineCount > exception.maxLines) {
      return {
        level: "error",
        message: `${relativePath}: ${lineCount} lines exceeds ratchet cap of ${exception.maxLines} (${exception.reason}). Refactor before adding more code.`,
      };
    }
    const utilization = lineCount / exception.maxLines;
    const tightenHint =
      utilization < 0.8
        ? `Cap is ${Math.round(utilization * 100)}% utilized — consider tightening to ${Math.ceil((lineCount + 1) / 50) * 50}.`
        : null;
    return {
      level: "info",
      message: `${relativePath}: ${lineCount}/${exception.maxLines} lines — ${exception.reason}`,
      tightenHint,
    };
  }

  if (lineCount > errorThreshold) {
    return {
      level: "error",
      message: `${relativePath}: ${lineCount} lines (max ${errorThreshold}).`,
      guidance: getFolderGuidance(relativePath),
    };
  }
  if (lineCount > warnThreshold) {
    return {
      level: "warn",
      message: `${relativePath}: ${lineCount} lines (warn at ${warnThreshold}).`,
      guidance: getFolderGuidance(relativePath),
    };
  }
  return null;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

function printResult(log, icon, { message, guidance, tightenHint }) {
  log(`  ${icon}  ${message}`);
  if (guidance) log(`        → ${guidance}`);
  if (tightenHint) log(`        ↓ ${tightenHint}`);
}

function reportResults(infos, warnings, errors, totalCount) {
  if (infos.length > 0) {
    console.log("Complexity exceptions (known technical debt):\n");
    for (const i of infos) printResult(console.log, "ℹ️", i);
    console.log("");
  }

  if (warnings.length > 0) {
    console.warn("Complexity warnings:\n");
    for (const w of warnings) printResult(console.warn, "⚠️ ", w);
    console.warn("");
  }

  if (errors.length > 0) {
    console.error("Complexity errors:\n");
    for (const e of errors) printResult(console.error, "❌", e);
    console.error("");
    process.exitCode = 1;
  } else {
    console.log(`Complexity check passed for ${totalCount} source files.`);
  }

  console.log(
    "\n  Legacy code is technical debt to be paid off — not a constraint to work around.\n",
  );
}

async function main() {
  try {
    await stat(sourceDirectory);
  } catch {
    console.log("No src directory found; skipping complexity check.");
    return;
  }

  const files = await collectFiles(sourceDirectory);
  const managedFiles = files.filter(isManagedFile);
  const analyses = await Promise.all(managedFiles.map(analyzeFile));

  const infos = [];
  const warnings = [];
  const errors = [];

  for (const file of analyses) {
    const lineResult = checkLineCount(file);
    if (!lineResult) continue;
    if (lineResult.level === "error") errors.push(lineResult);
    else if (lineResult.level === "warn") warnings.push(lineResult);
    else infos.push(lineResult);
  }

  reportResults(infos, warnings, errors, analyses.length);
}

await main();
