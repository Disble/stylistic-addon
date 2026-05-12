import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";

const rootDirectory = process.cwd();
const sourceDirectory = path.join(rootDirectory, "src");
const sourceExtensions = [".ts", ".tsx"];
const wordGlobalForbiddenRoots = ["src/domain/", "src/infrastructure/", "src/taskpane/"];
const textSearchPrimitiveNames = new Set([
  "findWhitespaceInsensitiveSlice",
  "findUniqueLocatorSubstring",
  "findFirstAlphanumericOffset",
  "removeWhitespaceWithIndices",
  "normalizeChar",
]);

/** Matches test files excluded from architecture rails to preserve test doubles and harness setup. */
function isExcludedTestFile(relativeFilePath) {
  return (
    relativeFilePath.endsWith(".test.ts") ||
    relativeFilePath.endsWith(".test.tsx") ||
    relativeFilePath.endsWith(".spec.ts") ||
    relativeFilePath.endsWith(".spec.tsx") ||
    relativeFilePath.endsWith("TestHelper.ts")
  );
}

/** Returns a repository-relative path with POSIX separators for deterministic diagnostics. */
function relativePath(filePath) {
  return path.relative(rootDirectory, filePath).replace(/\\/g, "/");
}

/** Returns whether a path starts with one of the configured repository-relative prefixes. */
function startsWithAny(relativeFilePath, prefixes) {
  return prefixes.some((prefix) => relativeFilePath.startsWith(prefix));
}

/** Collects TypeScript source files under a directory. */
async function collectSourceFiles(directory) {
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
      files.push(...(await collectSourceFiles(entryPath)));
      continue;
    }
    if (!entry.isFile()) {
      continue;
    }
    if (entry.name.endsWith(".d.ts")) {
      continue;
    }
    if (sourceExtensions.some((extension) => entry.name.endsWith(extension))) {
      files.push(entryPath);
    }
  }
  return files;
}

/** Removes comments and string bodies enough for rails checks to avoid documentation false positives. */
function stripCommentsAndStrings(content) {
  let sanitized = "";
  let index = 0;

  while (index < content.length) {
    const current = content[index];
    const next = content[index + 1];

    if (current === "/" && next === "*") {
      index += 2;
      while (index < content.length && !(content[index] === "*" && content[index + 1] === "/")) {
        index += 1;
      }
      index += 2;
      continue;
    }

    if (current === "/" && next === "/") {
      index += 2;
      while (index < content.length && content[index] !== "\n") {
        index += 1;
      }
      continue;
    }

    if (current === '"' || current === "'" || current === "`") {
      const quote = current;
      sanitized += quote.repeat(2);
      index += 1;

      while (index < content.length) {
        const char = content[index];
        if (char === "\\") {
          index += 2;
          continue;
        }

        index += 1;
        if (char === quote) {
          break;
        }
      }
      continue;
    }

    sanitized += current;
    index += 1;
  }

  return sanitized;
}

/** Extracts static import module specifiers from source content. */
function extractImportSpecifiers(content) {
  const specifiers = [];

  for (const rawLine of content.split(/\r?\n/u)) {
    const line = rawLine.trim();
    if (!line.startsWith("import ")) {
      continue;
    }

    const fromIndex = line.lastIndexOf(" from ");
    if (fromIndex >= 0) {
      const specifier = readQuotedSpecifier(line.slice(fromIndex + 6));
      if (specifier) {
        specifiers.push(specifier);
      }
      continue;
    }

    const sideEffectSpecifier = readQuotedSpecifier(line.slice("import ".length));
    if (sideEffectSpecifier) {
      specifiers.push(sideEffectSpecifier);
    }
  }

  return specifiers;
}

/** Reads one quoted module specifier from the start of an import tail. */
function readQuotedSpecifier(fragment) {
  const trimmed = fragment.trim();
  const quote = trimmed[0];
  if (quote !== '"' && quote !== "'") {
    return null;
  }

  let specifier = "";
  for (let index = 1; index < trimmed.length; index += 1) {
    const current = trimmed[index];
    if (current === "\\") {
      const escaped = trimmed[index + 1];
      if (escaped === undefined) {
        break;
      }
      specifier += escaped;
      index += 1;
      continue;
    }

    if (current === quote) {
      return specifier;
    }

    specifier += current;
  }

  return null;
}

/** Returns true when an import path targets adapters from inside domain. */
function isDomainAdapterImport(specifier) {
  return /^\.\.\/(?:\.\.\/)*adapters(?:\/|$)/.test(specifier);
}

/** Returns true when a forbidden concrete text-search import appears in an adapter file. */
function isForbiddenAdapterTextSearchImport(specifier) {
  return (
    specifier === "./WordTextLocatorAdapter" ||
    specifier === "../word/WordTextLocatorAdapter" ||
    specifier === "../../word/WordTextLocatorAdapter" ||
    /^\.\.\/(?:\.\.\/)*(?:domain|core)\/text-search(?:\/|$)/.test(specifier)
  );
}

/** Returns true when taskpane imports text-search internals instead of a port/workflow surface. */
function isForbiddenTaskpaneTextSearchImport(specifier) {
  return /^\.\.\/(?:\.\.\/)*(?:adapters\/word\/(?:WordTextLocatorAdapter|resolution)|domain\/text-search|core\/text-search)(?:\/|$)/.test(
    specifier,
  );
}

/** Adds diagnostics for import-based architecture violations. */
function validateImports(relativeFilePath, content, errors) {
  const specifiers = extractImportSpecifiers(content);

  for (const specifier of specifiers) {
    if (relativeFilePath.startsWith("src/domain/") && isDomainAdapterImport(specifier)) {
      errors.push(
        `${relativeFilePath}: domain layer must not import from adapters. Depend on ports instead (${specifier}).`,
      );
    }

    if (
      relativeFilePath.startsWith("src/adapters/word/") &&
      !relativeFilePath.endsWith("src/adapters/word/WordTextLocatorAdapter.ts") &&
      !relativeFilePath.endsWith("src/adapters/word/WordTextLocatorContext.ts") &&
      isForbiddenAdapterTextSearchImport(specifier)
    ) {
      errors.push(
        `${relativeFilePath}: concrete text-search implementations must not be imported here. Depend on an authorized locator port or orchestrator instead (${specifier}).`,
      );
    }

    if (
      (relativeFilePath.startsWith("src/domain/") || relativeFilePath.startsWith("src/taskpane/")) &&
      isForbiddenTaskpaneTextSearchImport(specifier)
    ) {
      errors.push(
        `${relativeFilePath}: taskpane/domain layers must not import text-search internals directly (${specifier}).`,
      );
    }
  }
}

/** Adds diagnostics for direct Word global execution outside the Word adapter boundary. */
function validateWordGlobal(relativeFilePath, executableContent, errors) {
  if (!startsWithAny(relativeFilePath, wordGlobalForbiddenRoots)) {
    return;
  }

  const wordCallPattern = /\bWord\s*\.\s*([A-Za-z_$][\w$]*)\s*\(/g;
  for (const match of executableContent.matchAll(wordCallPattern)) {
    errors.push(
      `${relativeFilePath}: Word global call Word.${match[1]}(...) is forbidden outside src/adapters/word/. Use IDocumentPort instead.`,
    );
  }
}

/** Adds diagnostics when canonical text-search primitives are redefined outside approved modules. */
function validateSearchPrimitiveRedefinition(relativeFilePath, executableContent, errors) {
  if (relativeFilePath.startsWith("src/domain/text-search/") || relativeFilePath.startsWith("src/core/text-search/")) {
    return;
  }

  const functionPattern = /\bfunction\s+([A-Za-z_$][\w$]*)\s*\(/g;
  for (const match of executableContent.matchAll(functionPattern)) {
    if (textSearchPrimitiveNames.has(match[1])) {
      errors.push(
        `${relativeFilePath}: canonical text-search primitive ${match[1]} must live only in the approved text-search core module.`,
      );
    }
  }
}

/** Validates one source file against architecture rails previously enforced by Biome GritQL. */
async function validateFile(filePath, errors) {
  const relativeFilePath = relativePath(filePath);
  if (isExcludedTestFile(relativeFilePath)) {
    return;
  }

  const content = await readFile(filePath, "utf8");
  const executableContent = stripCommentsAndStrings(content);

  validateImports(relativeFilePath, content, errors);
  validateWordGlobal(relativeFilePath, executableContent, errors);
  validateSearchPrimitiveRedefinition(relativeFilePath, executableContent, errors);
}

/** Runs deterministic architecture rails independent from formatter choice. */
async function main() {
  try {
    await stat(sourceDirectory);
  } catch {
    console.log("No src directory found; skipping architecture rails check.");
    return;
  }

  const files = await collectSourceFiles(sourceDirectory);
  const errors = [];
  await Promise.all(files.map((filePath) => validateFile(filePath, errors)));

  if (errors.length > 0) {
    console.error("Architecture rails check failed:\n");
    for (const error of errors) {
      console.error(`- ${error}`);
    }
    process.exitCode = 1;
    return;
  }

  console.log(`Architecture rails check passed for ${files.length} source files.`);
}

await main();
