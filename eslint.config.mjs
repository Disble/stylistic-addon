import globals from "globals";
import checkFile from "eslint-plugin-check-file";
import importX from "eslint-plugin-import-x";
import jsdoc from "eslint-plugin-jsdoc";
import officeAddins from "eslint-plugin-office-addins";
import reactDoctor from "react-doctor/eslint-plugin";
import reactPlugin from "eslint-plugin-react";
import sonarjs from "eslint-plugin-sonarjs";
import tseslint from "typescript-eslint";
import vitest from "eslint-plugin-vitest";

const vitestGlobals = {
  afterAll: "readonly",
  afterEach: "readonly",
  beforeAll: "readonly",
  beforeEach: "readonly",
  describe: "readonly",
  expect: "readonly",
  it: "readonly",
  test: "readonly",
  vi: "readonly",
};

export default [
  {
    ignores: ["coverage/**", "dist/**"],
  },
  ...officeAddins.configs.recommended,
  importX.configs["flat/recommended"],
  importX.configs["flat/typescript"],
  reactPlugin.configs.flat.recommended,
  reactPlugin.configs.flat["jsx-runtime"],
  {
    ...reactDoctor.configs.recommended,
    files: ["src/taskpane/**/*.{ts,tsx,js,jsx}"],
  },
  {
    files: ["**/*.{ts,tsx,js,jsx,mjs,cjs}"],
    plugins: {
      "@typescript-eslint": tseslint.plugin,
      jsdoc,
      "office-addins": officeAddins,
      react: reactPlugin,
      sonarjs,
    },
    languageOptions: {
      globals: {
        ...globals.browser,
        ...globals.es2021,
        ...globals.node,
        Office: "readonly",
        Word: "readonly",
      },
      parser: tseslint.parser,
    },
    settings: {
      react: {
        version: "detect",
      },
      "import-x/resolver-next": [],
    },
    rules: {
      "no-redeclare": "off",
      "import-x/no-duplicates": "error",
      "import-x/no-cycle": ["error", { maxDepth: 1 }],
      "import-x/no-unresolved": "off",
      "sonarjs/cognitive-complexity": ["warn", 15],
      "sonarjs/no-all-duplicated-branches": "warn",
      "sonarjs/no-identical-functions": "warn",
      "sonarjs/no-redundant-boolean": "warn",
      "sonarjs/no-small-switch": "warn",
      "jsdoc/check-tag-names": "error",
      "jsdoc/require-description": "error",
      "jsdoc/require-jsdoc": "off",
      "jsdoc/require-param": "off",
      "jsdoc/require-param-description": "off",
      "jsdoc/require-returns": "off",
      "jsdoc/require-returns-description": "off",
      "@typescript-eslint/no-unused-vars": [
        "error",
        {
          argsIgnorePattern: "^_",
          caughtErrorsIgnorePattern: "^_",
          varsIgnorePattern: "^_",
        },
      ],
    },
  },
  {
    files: ["src/**/*.types.ts"],
    ignores: ["src/**/*.test.ts", "src/**/*.test.tsx", "src/**/__tests__/**/*.{ts,tsx}"],
    rules: {
      "jsdoc/require-jsdoc": [
        "error",
        {
          contexts: ["TSInterfaceDeclaration", "TSTypeAliasDeclaration", "TSEnumDeclaration"],
          publicOnly: false,
        },
      ],
    },
  },
  {
    files: ["src/**/*.{ts,tsx}"],
    ignores: [
      "src/**/*.types.ts",
      "src/**/*.test.ts",
      "src/**/*.test.tsx",
      "src/**/__tests__/**/*.{ts,tsx}",
    ],
    rules: {
      "jsdoc/require-jsdoc": [
        "error",
        {
          publicOnly: true,
          require: {
            ArrowFunctionExpression: false,
            ClassDeclaration: true,
            ClassExpression: false,
            FunctionDeclaration: true,
            FunctionExpression: false,
            MethodDefinition: false,
          },
        },
      ],
    },
  },
  {
    files: ["src/**/*.{ts,tsx}"],
    ignores: [
      "src/**/*.types.ts",
      "src/**/*.constants.ts",
      "src/**/*.helpers.ts",
      "src/**/*TestHelper.ts",
      "src/**/*.test.ts",
      "src/**/*.test.tsx",
      "src/**/__tests__/**/*.{ts,tsx}",
      "src/domain/ports.ts",
      "src/infrastructure/config.ts",
      "src/taskpane/taskpane.ts",
      "src/**/index.ts",
      "src/**/index.tsx",
    ],
    rules: {
      "no-restricted-syntax": [
        "error",
        {
          selector:
            "Program:has(TSInterfaceDeclaration, TSTypeAliasDeclaration, TSEnumDeclaration):has(ClassDeclaration, FunctionDeclaration, VariableDeclaration)",
          message:
            "Implementation modules must not declare types/interfaces/enums. Move those contracts to a sibling *.types.ts file.",
        },
        {
          selector: "ExportNamedDeclaration[source]",
          message:
            "Runtime modules must not re-export from sibling files. Import the dedicated module directly instead of routing through implementation files.",
        },
        {
          selector:
            "ExportNamedDeclaration > VariableDeclaration[kind='const'] > VariableDeclarator[id.name=/^[A-Z0-9_]+$/]",
          message:
            "Exported SCREAMING_CASE constants belong in *.constants.ts (or infrastructure/config.ts for shared configuration).",
        },
        {
          selector:
            "Program:has(ExportNamedDeclaration > ClassDeclaration, ExportNamedDeclaration > FunctionDeclaration) > VariableDeclaration[kind='const']",
          message:
            "Runtime modules must not declare top-level constants. Move them to a sibling *.constants.ts file.",
        },
        {
          selector:
            "Program:has(> ExportNamedDeclaration > ClassDeclaration) > FunctionDeclaration",
          message:
            "Class implementation files must not declare top-level helpers. Move them to a sibling *.helpers.ts file.",
        },
        {
          selector:
            "Program:has(> ExportNamedDeclaration > ClassDeclaration) > ExportNamedDeclaration:has(> FunctionDeclaration)",
          message:
            "Class implementation files must not export top-level helpers. Move them to a sibling *.helpers.ts file.",
        },
      ],
    },
  },
  {
    files: ["src/**/*.constants.ts"],
    ignores: ["src/**/*.test.ts", "src/**/*.test.tsx", "src/**/__tests__/**/*.{ts,tsx}"],
    rules: {
      "no-restricted-syntax": [
        "error",
        {
          selector:
            "Program:has(TSInterfaceDeclaration, TSTypeAliasDeclaration, TSEnumDeclaration)",
          message:
            "Type/interface/enum declarations belong in a sibling *.types.ts file, not in constants modules.",
        },
      ],
    },
  },
  {
    files: ["src/taskpane/components/**/*.tsx"],
    rules: {
      "no-restricted-syntax": [
        "error",
        {
          selector:
            "Program > TSTypeAliasDeclaration, Program > TSInterfaceDeclaration, Program > TSEnumDeclaration",
          message:
            "Component .tsx files must not declare types/interfaces/enums. Move them to ComponentName.types.ts.",
        },
        {
          selector:
            "Program > ExportNamedDeclaration:has(> TSTypeAliasDeclaration, > TSInterfaceDeclaration, > TSEnumDeclaration)",
          message:
            "Component .tsx files must not export types/interfaces/enums. Move them to ComponentName.types.ts.",
        },
        {
          selector: "Program > VariableDeclaration[kind='const']",
          message:
            "Component .tsx files must not declare top-level constants. Move them to ComponentName.constants.ts.",
        },
        {
          selector: "Program > ExportNamedDeclaration:has(> VariableDeclaration[kind='const'])",
          message:
            "Component .tsx files must not export top-level constants. Move them to ComponentName.constants.ts.",
        },
        {
          selector: "Program > FunctionDeclaration",
          message:
            "Component .tsx files must not declare top-level helpers. Move them to ComponentName.helpers.ts or inline them inside the component when truly local.",
        },
        {
          selector:
            "Program > ExportNamedDeclaration:has(> FunctionDeclaration[id.name!=/^[A-Z]/])",
          message:
            "Component .tsx files must not export top-level helpers. Keep only the component export in ComponentName.tsx.",
        },
        {
          selector: "ImportSpecifier[imported.name='makeStyles']",
          message:
            "makeStyles must be declared in ComponentName.styles.ts and consumed from hooks/components through that sibling module.",
        },
        {
          selector: "CallExpression[callee.name='makeStyles']",
          message:
            "makeStyles calls must live in ComponentName.styles.ts, never inside component or hook files.",
        },
      ],
    },
  },
  {
    files: ["src/taskpane/components/**/*.hooks.ts"],
    rules: {
      "no-restricted-syntax": [
        "error",
        {
          selector:
            "Program > TSTypeAliasDeclaration, Program > TSInterfaceDeclaration, Program > TSEnumDeclaration",
          message:
            "Component hook files must not declare types/interfaces/enums. Move them to ComponentName.types.ts.",
        },
        {
          selector:
            "Program > ExportNamedDeclaration:has(> TSTypeAliasDeclaration, > TSInterfaceDeclaration, > TSEnumDeclaration)",
          message:
            "Component hook files must not export types/interfaces/enums. Move them to ComponentName.types.ts.",
        },
        {
          selector: "Program > VariableDeclaration[kind='const']",
          message:
            "Component hook files must not declare top-level constants. Move them to ComponentName.constants.ts.",
        },
        {
          selector: "Program > ExportNamedDeclaration:has(> VariableDeclaration[kind='const'])",
          message:
            "Component hook files must not export top-level constants. Move them to ComponentName.constants.ts.",
        },
        {
          selector: "Program > FunctionDeclaration[id.name!=/^use[A-Z]/]",
          message:
            "Component hook files must not declare non-hook top-level helpers. Move them to ComponentName.helpers.ts.",
        },
        {
          selector:
            "Program > ExportNamedDeclaration:has(> FunctionDeclaration[id.name!=/^use[A-Z]/])",
          message:
            "Component hook files must not export non-hook helpers. Move them to ComponentName.helpers.ts.",
        },
        {
          selector: "ImportSpecifier[imported.name='makeStyles']",
          message:
            "makeStyles must be declared in ComponentName.styles.ts and consumed from hooks/components through that sibling module.",
        },
        {
          selector: "CallExpression[callee.name='makeStyles']",
          message:
            "makeStyles calls must live in ComponentName.styles.ts, never inside component or hook files.",
        },
      ],
    },
  },
  {
    files: ["src/taskpane/components/**/*.helpers.ts"],
    rules: {
      "no-restricted-syntax": [
        "error",
        {
          selector:
            "Program > TSTypeAliasDeclaration, Program > TSInterfaceDeclaration, Program > TSEnumDeclaration",
          message:
            "Component helper files must not declare types/interfaces/enums. Move them to ComponentName.types.ts.",
        },
        {
          selector:
            "Program > ExportNamedDeclaration:has(> TSTypeAliasDeclaration, > TSInterfaceDeclaration, > TSEnumDeclaration)",
          message:
            "Component helper files must not export types/interfaces/enums. Move them to ComponentName.types.ts.",
        },
        {
          selector: "Program > VariableDeclaration[kind='const']",
          message:
            "Component helper files must not declare top-level constants. Move them to ComponentName.constants.ts.",
        },
        {
          selector: "Program > ExportNamedDeclaration:has(> VariableDeclaration[kind='const'])",
          message:
            "Component helper files must not export top-level constants. Move them to ComponentName.constants.ts.",
        },
      ],
    },
  },
  {
    files: ["src/**/*.{ts,tsx}"],
    plugins: {
      "check-file": checkFile,
    },
    rules: {
      "check-file/filename-blocklist": [
        "error",
        {
          "src/taskpane/components/**/use*.ts": "*.hooks.ts",
          "src/taskpane/components/**/use*.tsx": "*.tsx",
          "src/**/utils.ts": "*.ts",
          "src/**/Utils.ts": "*.ts",
        },
      ],
      "check-file/folder-match-with-fex": [
        "error",
        {
          "src/**/*.test.ts": "**/__tests__/",
          "src/**/*.test.tsx": "**/__tests__/",
        },
      ],
      "check-file/folder-naming-convention": [
        "error",
        {
          "src/taskpane/components/*/": "PASCAL_CASE",
        },
      ],
    },
  },
  {
    files: ["**/*.test.ts", "**/*.test.tsx", "**/*.spec.ts", "**/*.spec.tsx", "**/*TestHelper.ts"],
    languageOptions: {
      globals: vitestGlobals,
    },
    plugins: {
      vitest,
    },
    rules: {
      ...vitest.configs.recommended.rules,
      "vitest/expect-expect": "off",
      "vitest/no-identical-title": "off",
    },
  },
];
