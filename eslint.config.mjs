import globals from "globals";
import officeAddins from "eslint-plugin-office-addins";
import reactPlugin from "eslint-plugin-react";
import tseslint from "typescript-eslint";

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
  ...officeAddins.configs.recommended,
  reactPlugin.configs.flat.recommended,
  reactPlugin.configs.flat["jsx-runtime"],
  {
    files: ["**/*.{ts,tsx,js,jsx,mjs,cjs}"],
    plugins: {
      "@typescript-eslint": tseslint.plugin,
      "office-addins": officeAddins,
      react: reactPlugin,
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
    },
    rules: {
      "no-redeclare": "off",
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
    files: [
      "**/*.test.ts",
      "**/*.test.tsx",
      "**/*.spec.ts",
      "**/*.spec.tsx",
      "**/*TestHelper.ts",
    ],
    languageOptions: {
      globals: vitestGlobals,
    },
  },
];
