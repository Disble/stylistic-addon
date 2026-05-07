# Linting and File Anatomy Guidelines

This document defines the project rules for ESLint-driven file anatomy. The goal
is to stop generated or hand-written code from accumulating unrelated concerns in
one file.

These rules deliberately avoid vague categories like “main file” or “architectural
file”. A rule is valid only when it can be detected mechanically by filename,
folder path, import, or AST shape.

## Guiding principle

When code contains a recognizable secondary concern, that concern must live in
its dedicated anatomy file.

| Detected concern | Detector | Required file |
| --- | --- | --- |
| types, interfaces, enums | `type`, `interface`, `enum` declarations | `*.types.ts` |
| Fluent UI/Griffel styles | `makeStyles(...)` | `*.styles.ts` |
| exported semantic constants | `export const SCREAMING_CASE` | `*.constants.ts` |
| pure exported helpers | exported non-hook helper functions | `*.helpers.ts` |
| schemas | schema declarations/exports | `*.schema.ts` |
| tests | `*.test.ts`, `*.test.tsx` | sibling `__tests__/` directory anywhere under `src/` |
| test helpers and mocks | `*TestHelper.ts`, `*Mocks.ts` | test-support files or explicit test helper modules |

The enforcement model is content-first: if a new module role appears tomorrow,
the rules still hold because they target the mixed concern itself, not a brittle
taxonomy of possible file roles.

## React component anatomy

Every component under `src/taskpane/components/` uses a PascalCase folder and
basename-aligned anatomy files.

```txt
ComponentName/
├─ ComponentName.tsx
├─ ComponentName.types.ts
├─ ComponentName.styles.ts
├─ ComponentName.hooks.ts
├─ ComponentName.constants.ts
├─ ComponentName.helpers.ts
├─ ComponentName.schema.ts
├─ index.ts
└─ __tests__/
   └─ ComponentName.test.tsx
```

Required files:

- `ComponentName.tsx`
- `ComponentName.types.ts`
- `index.ts`

Optional files are added only when their concern exists. Do not create empty
anatomy files.

### Component file rules

`ComponentName.tsx` contains the presentational component only.

It must not declare or export:

- types, interfaces, or enums,
- `makeStyles(...)`,
- semantic constants,
- hooks,
- helpers,
- secondary React components.

It must not declare top-level constants or top-level helper functions either.
If a presentational component needs either of those concerns, extract them to
`ComponentName.constants.ts` or `ComponentName.helpers.ts`.

### Styles file rules

Any call to `makeStyles(...)` belongs in `ComponentName.styles.ts`.

```ts
// ComponentName.styles.ts
export const useComponentNameStyles = makeStyles({
  root: {},
});
```

No `makeStyles(...)` call should appear in `ComponentName.tsx`,
`ComponentName.hooks.ts`, or any other non-`.styles.ts` file.

### Hooks file rules

Component-owned hooks live in `ComponentName.hooks.ts`, not in
`useComponentName.ts`.

```ts
// ComponentName.hooks.ts
export function useComponentName(): ComponentNameClasses {
  const styles = useComponentNameStyles();
  return { root: styles.root };
}
```

This keeps every component-owned file aligned to the component basename and gives
ESLint/check-file one deterministic folder contract to enforce.

`ComponentName.hooks.ts` must not declare or export:

- types, interfaces, or enums,
- semantic constants,
- non-hook helper functions.

If a helper does not start with `use`, it does not belong in the hooks file.
Move it to `ComponentName.helpers.ts`.

### Helpers file rules

`ComponentName.helpers.ts` is for pure exported helper functions only.

It must not declare or export:

- types, interfaces, or enums,
- semantic constants.

If helpers need named constants, extract them to `ComponentName.constants.ts` and
import them from there.

## General TypeScript anatomy

Outside React component folders, the same content-first rules apply.

### Tests

All tests under `src/` must live in a sibling `__tests__/` directory.

```txt
src/domain/pipeline/PipelineOrchestrator.ts
src/domain/pipeline/__tests__/PipelineOrchestrator.test.ts

src/adapters/word/WordAdapter.ts
src/adapters/word/__tests__/WordAdapterAcceptSuggestion.test.ts
```

This is enforced mechanically through `eslint-plugin-check-file` using
`folder-match-with-fex`, so `src/foo/Bar.test.ts` is invalid while
`src/foo/__tests__/Bar.test.ts` is valid.

### Types

TypeScript contracts live in `*.types.ts`, whether they are exported or local to
one implementation module. Local DTO shapes and private normalization contracts
are still secondary concerns; leaving them in implementation files is a mixed
responsibility false negative.

```txt
TaskpaneAuthStore.types.ts
ResultsPanelFilters.types.ts
BetterAuthAdapter.types.ts
```

### Constants

Exported semantic constants live in `*.constants.ts`.

```txt
TaskpaneAuthStore.constants.ts
ApplicationConfig.constants.ts
```

Existing legacy modules such as `src/infrastructure/config.ts` may be migrated
gradually, but new constants should use the `.constants.ts` suffix unless a
documented exception exists.

### Helpers

Exported pure helper functions live in `*.helpers.ts`.

Helpers must not import React, Zustand, Fluent UI, Office.js, Mastra clients, or
other framework/host APIs. If a function needs those dependencies, it is not a
generic helper; it belongs in a more explicit module.

### Runtime implementation files

Files whose runtime export is a class or function must not also declare top-level
constants, and class files must not also declare top-level helper functions or
sibling-file re-exports.

```txt
BatchApplyOrchestrator.ts              ✅ class only
BatchApplyOrchestrator.types.ts        ✅ contracts
BatchApplyOrchestrator.constants.ts    ✅ semantic constants when needed
BatchApplyOrchestrator.helpers.ts      ✅ pure top-level helpers when needed
chunker.ts                             ✅ exported runtime function only
chunker.constants.ts                   ✅ semantic constants when needed
```

This keeps class modules deterministic: the class file owns orchestration or
behavior, while every secondary concern lives in a dedicated sibling file.

Narrow, explicit exceptions:

- `src/taskpane/taskpane.ts` is the composition root and may keep module-level
  collaborator wiring.
- `*TestHelper.ts` files are test-support modules and may keep fixture/runtime
  setup values without forcing production anatomy rules onto test scaffolding.

Exported runtime singletons such as Zustand stores are not treated as semantic
constants. The rule targets auxiliary top-level constants, not the module's main
runtime export.

### Stores

Store files such as `*Store.ts` should contain the store runtime and store actions
only. Store state types and initial state constants should be extracted:

```txt
TaskpaneAuthStore.ts
TaskpaneAuthStore.types.ts
TaskpaneAuthStore.constants.ts
```

## ESLint strategy

Prefer established ESLint plugins before adding custom scripts.

Planned plugin responsibilities:

- `eslint-plugin-check-file`: filename and folder contracts.
- `eslint-plugin-boundaries`: layer and dependency boundaries.
- `eslint-plugin-sonarjs`: Sonar-aligned maintainability rules.
- `eslint-plugin-import-x`: import/export hygiene.
- `eslint-plugin-vitest`: Vitest-specific rules.
- `eslint-plugin-jsdoc`: documentation comments for public functions/classes.

`eslint-plugin-unicorn` is intentionally not part of the initial plan because the
project is expected to align with Sonar feedback and Unicorn can add overlapping,
opinionated diagnostics.

Custom scripts remain acceptable only for repository-specific checks that cannot
be expressed cleanly through ESLint.
