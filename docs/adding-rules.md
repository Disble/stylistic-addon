# Adding Rules

This guide explains how to add new editorial rules to Stylistic's analysis engine. No framework knowledge is required — rules are declared as data, not code.

## How Rules Work

Every rule in Stylistic is an `AnalysisRule` object with three properties:

```typescript
interface AnalysisRule {
  id: string;                        // Stable ID, used as suggestion prefix
  category: string;                  // Human-readable label for justifications
  detect(text: string): Suggestion[];  // Scan text, return suggestions
}
```

You don't implement `AnalysisRule` manually. Instead, you declare an array of patterns and use the `buildPatternRule()` factory:

```typescript
const myRule = buildPatternRule("my-rule-id", "Category Name", [
  {
    pattern: /some phrase/gi,
    replacement: "better phrase",
    justification: "Explanation shown to the user.",
  },
]);
```

## Step-by-Step

### 1. Define your patterns

Open `src/lib/analyzer.ts`. Before the `rules` array, add your rule definition:

```typescript
const passiveVoiceRule = buildPatternRule("passive", "Voz pasiva", [
  {
    pattern: /fue realizado/gi,
    replacement: "se realizó",
    justification: "La voz activa es más directa y clara.",
  },
  {
    pattern: /es requerido/gi,
    replacement: "se requiere",
    justification: "La voz activa es más directa y clara.",
  },
]);
```

### 2. Register the rule

Add your rule to the `rules` array:

```typescript
const rules: AnalysisRule[] = [
  redundancyRule,
  fillerRule,
  wordChoiceRule,
  passiveVoiceRule,  // ← add here
];
```

### 3. Build and test

```bash
npm run build
npm start
```

Open Word, paste text containing your patterns, and click "Analizar y sugerir".

That's it. No other files need to change.

## Pattern Reference

### PatternEntry fields

| Field | Type | Description |
|---|---|---|
| `pattern` | `RegExp` | Regex to match. Use `g` for global, `i` for case-insensitive. |
| `replacement` | `string` | Text to replace the match with. Use `""` to delete. |
| `justification` | `string` | User-facing explanation. Shown in the results panel. |

### Tips for writing patterns

**Use the `gi` flags.** Global (`g`) finds all occurrences. Case-insensitive (`i`) catches "Utilizar" and "utilizar".

```typescript
pattern: /utilizar/gi
```

**Match optional punctuation.** Fillers often have a trailing comma and space:

```typescript
pattern: /en realidad,?\s*/gi
replacement: ""
```

**Use `""` for deletions.** When the replacement is an empty string, the matched text is simply removed:

```typescript
{
  pattern: /obviamente,?\s*/gi,
  replacement: "",
  justification: "If it's obvious, it doesn't need saying.",
}
```

**Avoid overly broad patterns.** A pattern like `/el/gi` would match inside words like "modelo". Prefer specific multi-word phrases or use word boundaries:

```typescript
pattern: /\brealizar\b/gi
```

**Keep justifications concise.** They appear in a compact UI. One sentence is ideal.

### Choosing a rule ID

The `id` is used as a prefix for suggestion IDs (e.g., `"passive-0"`, `"passive-1"`). Use:

- Lowercase, no spaces
- Descriptive of the category
- Unique across all rules

### Choosing a category name

The `category` is prepended to justification text in brackets (e.g., `[Voz pasiva] La voz activa es más directa.`). Use a short, human-readable label in the language of the target audience.

## Deduplication Behavior

If two rules flag the same text (e.g., *"en realidad"* is both a filler and a redundancy), only the first match is kept. The deduplication is case-insensitive and based on `originalText`.

This means the order of rules in the `rules` array determines priority. Rules listed first take precedence.

## Testing Rules Without Word

Since `analyzer.ts` has zero dependencies on Office.js, you can test rules in isolation:

```typescript
import { analyze } from "./src/lib/analyzer";

const suggestions = analyze(
  "Básicamente, es completamente necesario utilizar este periodo de tiempo."
);

console.log(suggestions);
// [
//   { id: "filler-0", originalText: "Básicamente, ", suggestedText: "", ... },
//   { id: "redundancy-0", originalText: "completamente necesario", suggestedText: "necesario", ... },
//   { id: "wordchoice-0", originalText: "utilizar", suggestedText: "usar", ... },
//   { id: "redundancy-1", originalText: "periodo de tiempo", suggestedText: "periodo", ... },
// ]
```

This makes it straightforward to add unit tests for new rules using any test framework (Jest, Vitest, etc.).
