import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";

// ---------------------------------------------------------------------------
// Thresholds
// ---------------------------------------------------------------------------

const LINE_WARN = 500;
const LINE_ERROR = 800;

/**
 * Known exceptions — files whose complexity is inherent, not a design smell.
 * Each entry acts as a RATCHET: the allowed max should DECREASE over time
 * as refactors chip away at the file. Raising a cap requires justification.
 */
const KNOWN_EXCEPTIONS = new Map([
  [
    "src/adapters/word/ResolveSuggestionCommand.ts",
    {
      maxLines: 1100,
      reason: "Inherent Word tracked-change resolution complexity",
    },
  ],
]);

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
    if (entry.name.endsWith(".test.ts") || entry.name.endsWith(".spec.ts"))
      continue;
    if (entry.name.includes("TestHelper")) continue;
    if (supportedExtensions.some((ext) => entry.name.endsWith(ext))) {
      files.push(entryPath);
    }
  }
  return files;
}

function isManagedFile(filePath) {
  return managedRoots.some((root) => filePath.startsWith(`${root}${path.sep}`));
}

// ---------------------------------------------------------------------------
// Metrics
// ---------------------------------------------------------------------------

const CONTROL_FLOW_KEYWORDS = new Set([
  "if",
  "for",
  "while",
  "switch",
  "return",
  "const",
  "let",
  "var",
  "throw",
  "await",
  "try",
  "catch",
  "else",
  "new",
  "typeof",
  "import",
  "export",
  "from",
  "type",
  "interface",
  "enum",
  "class",
  "super",
  "yield",
  "delete",
  "void",
  "break",
  "continue",
  "case",
  "default",
]);

/**
 * Counts class methods in a TypeScript source file.
 * Heuristic: lines at class-body indent (2+ spaces) matching a method
 * declaration pattern, excluding control-flow keywords and constructors.
 */
function countClassMethods(content) {
  const lines = content.split("\n");
  let insideClass = false;
  let braceDepth = 0;
  let classStartDepth = 0;
  let count = 0;

  for (const line of lines) {
    // Detect class declaration
    if (!insideClass && /^\s*(?:export\s+)?class\s+/.test(line)) {
      insideClass = true;
      classStartDepth = braceDepth;
    }

    // Track brace depth (ignoring strings/comments — good enough for a guard)
    for (const ch of line) {
      if (ch === "{") braceDepth++;
      if (ch === "}") {
        braceDepth--;
        if (insideClass && braceDepth <= classStartDepth) {
          insideClass = false;
        }
      }
    }

    if (!insideClass) continue;

    // Match method declarations at class body level
    const match = line.match(
      /^\s{2}(async\s+)?(private\s+|protected\s+|public\s+|readonly\s+|static\s+)*([a-zA-Z_]\w*)\s*[<(]/,
    );
    if (!match) continue;

    const methodName = match[3];
    if (CONTROL_FLOW_KEYWORDS.has(methodName)) continue;
    if (methodName === "constructor") continue;

    count++;
  }

  return count;
}

/**
 * Counts module-level function declarations (not inside classes).
 */
function countModuleFunctions(content) {
  const lines = content.split("\n");
  let count = 0;

  for (const line of lines) {
    // export function foo( / function foo( / async function foo(
    if (
      /^(?:export\s+)?(?:async\s+)?function\s+[a-zA-Z_]\w*\s*[<(]/.test(line)
    ) {
      count++;
      continue;
    }
    // export const foo = ( / const foo = async (
    if (
      /^(?:export\s+)?const\s+[a-zA-Z_]\w*\s*=\s*(?:async\s+)?[(<]/.test(line)
    ) {
      count++;
    }
  }

  return count;
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
  const classMethods = countClassMethods(content);
  const moduleFunctions = countModuleFunctions(content);

  return { relativePath, lineCount, classMethods, moduleFunctions };
}

function checkLineCount(file) {
  const { relativePath, lineCount } = file;

  // Check exemptions
  if (EXEMPT_PATTERNS.some((pattern) => pattern.test(relativePath))) {
    return null;
  }

  // Check known exceptions (ratchet)
  const exception = KNOWN_EXCEPTIONS.get(relativePath);
  if (exception) {
    if (lineCount > exception.maxLines) {
      return {
        level: "error",
        message: `${relativePath}: ${lineCount} lines exceeds ratchet cap of ${exception.maxLines} (${exception.reason}). Refactor before adding more code.`,
      };
    }
    return null;
  }

  // Standard thresholds
  if (lineCount > LINE_ERROR) {
    return {
      level: "error",
      message: `${relativePath}: ${lineCount} lines (max ${LINE_ERROR}). Extract responsibilities into focused modules.`,
    };
  }
  if (lineCount > LINE_WARN) {
    return {
      level: "warn",
      message: `${relativePath}: ${lineCount} lines (warn at ${LINE_WARN}). Consider extracting before it grows further.`,
    };
  }
  return null;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

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

  const warnings = [];
  const errors = [];

  for (const file of analyses) {
    const lineResult = checkLineCount(file);
    if (lineResult) {
      if (lineResult.level === "error") errors.push(lineResult.message);
      else warnings.push(lineResult.message);
    }
  }

  // Report
  if (warnings.length > 0) {
    console.warn("Complexity warnings:\n");
    for (const w of warnings) console.warn(`  ⚠️  ${w}`);
    console.warn("");
  }

  if (errors.length > 0) {
    console.error("Complexity errors:\n");
    for (const e of errors) console.error(`  ❌ ${e}`);
    console.error("");
    process.exitCode = 1;
    return;
  }

  console.log(`Complexity check passed for ${analyses.length} source files.`);
}

await main();
