import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";

const rootDirectory = process.cwd();
const taskpaneDirectory = path.join(rootDirectory, "src", "taskpane");
const componentsDirectory = path.join(taskpaneDirectory, "components");
const allowedRootFiles = new Set(["index.ts"]);
const allowedSupportingFileSuffixes = [
  ".types.ts",
  ".styles.ts",
  ".hooks.ts",
  ".constants.ts",
  ".helpers.ts",
  ".schema.ts",
];

/**
 * Returns true when a filesystem path exists.
 * Missing React component folders are valid during the tooling-first migration stage.
 */
async function pathExists(filePath) {
  try {
    await stat(filePath);
    return true;
  } catch {
    return false;
  }
}

/** Validates the strict PascalCase component folder contract. */
function isPascalCase(name) {
  return /^[A-Z][A-Za-z0-9]*$/.test(name);
}

/** Converts an absolute path into a stable repository-relative path for diagnostics. */
function relativePath(filePath) {
  return path.relative(rootDirectory, filePath).replace(/\\/g, "/");
}

/** Reads a directory and returns an empty list when it does not exist. */
async function readDirectory(directory) {
  try {
    return await readdir(directory, { withFileTypes: true });
  } catch {
    return [];
  }
}

/** Ensures root-level component files do not bypass the folder anatomy. */
function validateComponentsRoot(entries, errors) {
  for (const entry of entries) {
    if (!entry.isFile()) {
      continue;
    }

    if (allowedRootFiles.has(entry.name)) {
      continue;
    }

    errors.push(
      `${relativePath(path.join(componentsDirectory, entry.name))}: React components must live in their own PascalCase folder; flat component files are forbidden.`,
    );
  }
}

/** Validates that a component folder contains only approved anatomy files. */
function validateAllowedComponentFile(componentName, fileName, componentDirectory, errors) {
  const allowedFiles = new Set([
    "index.ts",
    `${componentName}.tsx`,
    `${componentName}.types.ts`,
    `${componentName}.styles.ts`,
    `${componentName}.hooks.ts`,
    `${componentName}.constants.ts`,
    `${componentName}.helpers.ts`,
    `${componentName}.schema.ts`,
  ]);

  if (allowedFiles.has(fileName)) {
    return;
  }

  errors.push(
    `${relativePath(path.join(componentDirectory, fileName))}: unsupported component anatomy file. Use ${componentName}.tsx, ${componentName}.types.ts, optional styles/hooks/constants/helpers/schema, index.ts, or __tests__/.`,
  );
}

/** Validates the required minimum anatomy for one component directory. */
async function validateRequiredFiles(componentName, componentDirectory, errors) {
  const requiredFiles = [
    "index.ts",
    `${componentName}.tsx`,
    `${componentName}.types.ts`,
  ];

  for (const fileName of requiredFiles) {
    const filePath = path.join(componentDirectory, fileName);
    if (!(await pathExists(filePath))) {
      errors.push(
        `${relativePath(componentDirectory)}: missing required React component file ${fileName}.`,
      );
    }
  }
}

/** Validates TSX content stays presentational instead of becoming a mixed-responsibility module. */
async function validateTsxContent(componentName, componentDirectory, errors) {
  const componentFile = path.join(componentDirectory, `${componentName}.tsx`);
  if (!(await pathExists(componentFile))) {
    return;
  }

  const content = await readFile(componentFile, "utf8");
  const diagnostics = [
    {
      pattern: /^\s*(export\s+)?(interface|type|enum)\s+/m,
      message: "types/interfaces/enums belong in ComponentName.types.ts, not in the TSX file.",
    },
    {
      pattern: /^\s*(export\s+)?const\s+[A-Z0-9_]+\s*=/m,
      message: "presentation constants belong in ComponentName.constants.ts, not in the TSX file.",
    },
    {
      pattern: /^\s*(export\s+)?function\s+use[A-Z]/m,
      message: "component hooks belong in ComponentName.hooks.ts, not in the TSX file.",
    },
    {
      pattern: /^\s*(export\s+)?const\s+use[A-Z][A-Za-z0-9]*\s*=/m,
      message: "component hooks belong in ComponentName.hooks.ts, not in the TSX file.",
    },
    {
      pattern: /\bmakeStyles\s*\(/m,
      message: "Fluent UI style factories belong in ComponentName.styles.ts, not in the TSX file.",
    },
  ];

  for (const { pattern, message } of diagnostics) {
    if (pattern.test(content)) {
      errors.push(`${relativePath(componentFile)}: ${message}`);
    }
  }

  const declaredComponents = [
    ...content.matchAll(/^\s*(?:export\s+)?function\s+([A-Z][A-Za-z0-9]*)\s*\(/gm),
    ...content.matchAll(/^\s*(?:export\s+)?const\s+([A-Z][A-Za-z0-9]*)\s*=/gm),
  ].map((match) => match[1]);

  const foreignComponents = declaredComponents.filter((name) => name !== componentName);
  if (foreignComponents.length > 0) {
    errors.push(
      `${relativePath(componentFile)}: only ${componentName} may be declared here. Move ${foreignComponents.join(", ")} into their own component folders.`,
    );
  }
}

/** Validates one PascalCase component folder and its strict anatomy. */
async function validateComponentDirectory(entry, errors) {
  const componentName = entry.name;
  const componentDirectory = path.join(componentsDirectory, componentName);

  if (!isPascalCase(componentName)) {
    errors.push(
      `${relativePath(componentDirectory)}: component directories must use PascalCase and match their component name.`,
    );
    return;
  }

  const entries = await readDirectory(componentDirectory);

  for (const child of entries) {
    if (child.isDirectory()) {
      if (child.name !== "__tests__") {
        errors.push(
          `${relativePath(path.join(componentDirectory, child.name))}: only __tests__/ is allowed as a nested directory inside a component folder.`,
        );
      }
      continue;
    }

    if (!child.isFile()) {
      continue;
    }

    validateAllowedComponentFile(componentName, child.name, componentDirectory, errors);

    const isSupportingFile = allowedSupportingFileSuffixes.some((suffix) =>
      child.name.endsWith(suffix),
    );
    if (isSupportingFile && !child.name.startsWith(`${componentName}.`)) {
      errors.push(
        `${relativePath(path.join(componentDirectory, child.name))}: supporting files must share the component basename ${componentName}.`,
      );
    }
  }

  await validateRequiredFiles(componentName, componentDirectory, errors);
  await validateTsxContent(componentName, componentDirectory, errors);
}

/** Runs strict component-structure checks designed for LLM-authored React code. */
async function main() {
  if (!(await pathExists(componentsDirectory))) {
    console.log("No src/taskpane/components directory found; skipping React component rails check.");
    return;
  }

  const entries = await readDirectory(componentsDirectory);
  const errors = [];

  validateComponentsRoot(entries, errors);

  for (const entry of entries) {
    if (entry.isDirectory()) {
      await validateComponentDirectory(entry, errors);
    }
  }

  if (errors.length > 0) {
    console.error("React component rails check failed:\n");
    for (const error of errors) {
      console.error(`- ${error}`);
    }
    process.exitCode = 1;
    return;
  }

  console.log("React component rails check passed.");
}

await main();
