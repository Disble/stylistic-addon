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

    if (
      supportedExtensions.some((extension) => entry.name.endsWith(extension))
    ) {
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
  return filePath.startsWith(
    `${path.join(sourceDirectory, "adapters")}${path.sep}`,
  );
}

function stripSourceExtension(fileName) {
  const extension = supportedExtensions.find((candidate) =>
    fileName.endsWith(candidate),
  );

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

    const parsed = stripSourceExtension(path.basename(filePath));

    if (!parsed) {
      continue;
    }

    const { stem } = parsed;
    const parts = stem.split(".");

    // Rule 1: forbid generic utils.ts / Utils.ts in managed architecture folders.
    if (parts[0] === "utils" || parts[0] === "Utils") {
      errors.push(
        `${path.relative(rootDirectory, filePath)}: avoid generic utils filenames in architecture modules; use a descriptive feature name instead.`,
      );
      continue;
    }

    // Rule 2: forbid triple-compound filenames (e.g. foo.bar.baz.ts).
    if (parts.length > 2) {
      errors.push(
        `${path.relative(rootDirectory, filePath)}: too many dot-separated segments; use at most one role suffix (.test, .spec) after the base class name.`,
      );
      continue;
    }

    // Test / spec files and dedicated test helpers pass all structural rules.
    if (isTestFile(stem) || isTestHelperFile(stem)) {
      continue;
    }

    // Rule 3: files inside any handlers/ directory must end with Handler.
    if (isInHandlersDirectory(filePath)) {
      if (!parts[0].endsWith("Handler")) {
        errors.push(
          `${path.relative(rootDirectory, filePath)}: files inside a handlers/ directory must end with Handler (e.g. ReadTextHandler.ts).`,
        );
      }
      continue;
    }

    // Rule 4: files inside src/adapters/ must use a known OOP suffix.
    // A leading Mock prefix is allowed (e.g. MockFeedbackAdapter.ts).
    if (isInAdaptersRoot(filePath)) {
      const baseName = parts[0];
      const nameWithoutMock = baseName.startsWith("Mock")
        ? baseName.slice(4)
        : baseName;
      const hasSuffix = [...knownAdapterSuffixes].some((suffix) =>
        nameWithoutMock.endsWith(suffix),
      );

      if (!hasSuffix) {
        errors.push(
          `${path.relative(rootDirectory, filePath)}: adapter files must end with a known architectural suffix (Adapter, Decorator, Command, Builder, Cleanup, Factory, Inspector, Locator, Observer, Parser, Resolver, Executor, Machine, Events, Context, Orchestrator). Found: ${baseName}`,
        );
      }
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
