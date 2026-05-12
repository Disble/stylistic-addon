import { readdir, stat } from "node:fs/promises";
import path from "node:path";

const rootDirectory = process.cwd();
const sourceDirectory = path.join(rootDirectory, "src");
const managedRoots = [
  path.join(sourceDirectory, "domain"),
  path.join(sourceDirectory, "adapters"),
  path.join(sourceDirectory, "infrastructure"),
];
const supportedExtensions = [".ts", ".tsx", ".js", ".mjs", ".cjs"];

/**
 * Known OOP suffixes required for files inside src/adapters/.
 * Every non-test file must end with one of these (after stripping a leading "Mock" prefix).
 */
const knownAdapterSuffixes = new Set([
  "Adapter",
  "Decorator",
  "Command",
  "Builder",
  "Cleanup",
  "Factory",
  "Inspector",
  "Locator",
  "Observer",
  "Parser",
  "Resolver",
  "Executor",
  "Machine",
  "Events",
  "Context",
  "Orchestrator",
]);

async function collectFiles(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];

  for (const entry of entries) {
    const entryPath = path.join(directory, entry.name);

    if (entry.isDirectory()) {
      files.push(...(await collectFiles(entryPath)));
      continue;
    }

    if (!entry.isFile()) {
      continue;
    }

    if (entry.name.endsWith(".d.ts")) {
      continue;
    }

    if (supportedExtensions.some((extension) => entry.name.endsWith(extension))) {
      files.push(entryPath);
    }
  }

  return files;
}

function isManagedFile(filePath) {
  return managedRoots.some((root) => filePath.startsWith(`${root}${path.sep}`));
}

function isInHandlersDirectory(filePath) {
  return filePath.split(path.sep).includes("handlers");
}

function isInAdaptersRoot(filePath) {
  return filePath.startsWith(`${path.join(sourceDirectory, "adapters")}${path.sep}`);
}

function stripSourceExtension(fileName) {
  const extension = supportedExtensions.find((candidate) => fileName.endsWith(candidate));

  if (!extension) {
    return null;
  }

  return {
    extension,
    stem: fileName.slice(0, -extension.length),
  };
}

function isTestFile(stem) {
  return stem.endsWith(".test") || stem.endsWith(".spec");
}

function isTestHelperFile(stem) {
  return stem.endsWith("TestHelper") || stem.includes("TestHelper.");
}

async function pathExists(filePath) {
  try {
    await stat(filePath);
    return true;
  } catch {
    return false;
  }
}

function createRuleError(filePath, message) {
  return `${path.relative(rootDirectory, filePath)}: ${message}`;
}

function isGenericUtilsFile(parts) {
  return parts[0] === "utils" || parts[0] === "Utils";
}

function hasTooManySegments(parts) {
  return parts.length > 2;
}

function resolveAdapterBaseName(parts) {
  const baseName = parts[0];
  return baseName.startsWith("Mock") ? baseName.slice(4) : baseName;
}

function hasKnownAdapterSuffix(baseName) {
  return [...knownAdapterSuffixes].some((suffix) => baseName.endsWith(suffix));
}

function validateGenericUtilsRule(filePath, parts) {
  if (!isGenericUtilsFile(parts)) {
    return undefined;
  }

  return createRuleError(
    filePath,
    "avoid generic utils filenames in architecture modules; use a descriptive feature name instead."
  );
}

function validateSegmentCountRule(filePath, parts) {
  if (!hasTooManySegments(parts)) {
    return undefined;
  }

  return createRuleError(
    filePath,
    "too many dot-separated segments; use at most one role suffix (.test, .spec) after the base class name."
  );
}

function validateHandlerSuffixRule(filePath, parts) {
  if (!isInHandlersDirectory(filePath) || parts[0].endsWith("Handler")) {
    return undefined;
  }

  return createRuleError(
    filePath,
    "files inside a handlers/ directory must end with Handler (e.g. ReadTextHandler.ts)."
  );
}

function validateAdapterSuffixRule(filePath, parts) {
  if (!isInAdaptersRoot(filePath)) {
    return undefined;
  }

  const baseName = parts[0];
  const nameWithoutMock = resolveAdapterBaseName(parts);
  if (hasKnownAdapterSuffix(nameWithoutMock)) {
    return undefined;
  }

  return createRuleError(
    filePath,
    `adapter files must end with a known architectural suffix (Adapter, Decorator, Command, Builder, Cleanup, Factory, Inspector, Locator, Observer, Parser, Resolver, Executor, Machine, Events, Context, Orchestrator). Found: ${baseName}`
  );
}

function validateManagedFile(filePath) {
  const parsed = stripSourceExtension(path.basename(filePath));
  if (!parsed) {
    return undefined;
  }

  const { stem } = parsed;
  const parts = stem.split(".");
  const genericUtilsError = validateGenericUtilsRule(filePath, parts);
  if (genericUtilsError) {
    return genericUtilsError;
  }

  const segmentCountError = validateSegmentCountRule(filePath, parts);
  if (segmentCountError) {
    return segmentCountError;
  }

  if (isTestFile(stem) || isTestHelperFile(stem)) {
    return undefined;
  }

  const handlerError = validateHandlerSuffixRule(filePath, parts);
  if (handlerError) {
    return handlerError;
  }

  return validateAdapterSuffixRule(filePath, parts);
}

async function main() {
  if (!(await pathExists(sourceDirectory))) {
    console.log("No src directory found; skipping file naming check.");
    return;
  }

  const files = await collectFiles(sourceDirectory);
  const errors = [];

  for (const filePath of files) {
    if (!isManagedFile(filePath)) {
      continue;
    }

    const validationError = validateManagedFile(filePath);
    if (validationError) {
      errors.push(validationError);
    }
  }

  if (errors.length > 0) {
    console.error("File naming check failed:\n");

    for (const error of errors) {
      console.error(`- ${error}`);
    }

    process.exitCode = 1;
    return;
  }

  console.log(`File naming check passed for ${files.length} source files.`);
}

await main();
