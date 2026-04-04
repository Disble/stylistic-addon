# Backend Contract Clarification — Replace Suggestions

This document clarifies how the Stylistic Word add-in expects **replace-style** suggestions to be shaped by the backend.

The goal is simple:

> A replace suggestion must describe **what exact substring in the input should be replaced** and **what exact replacement text should take its place**.

It must **not** send the entire rewritten sentence in `suggestedText` when only one substring is being replaced.

---

## Why this clarification is necessary

The frontend apply flow works with this mental model:

1. find the `context`,
2. find the `anchor` inside that context,
3. replace the `anchor` range with `suggestedText`.

That means the current contract is **substring replacement**, not **full-context rewrite**.

If the backend sends a full sentence in `suggestedText` while `anchor` contains only a fragment, the add-in will replace only that fragment and the document can end up duplicated or malformed.

---

## Current semantic contract

For a `track-change` suggestion:

- `context` = paragraph-level locator used to find the suggestion in the input text,
- `anchor` = exact substring inside `context` that should be replaced,
- `suggestedText` = exact text that should replace `anchor`.

This means:

- `suggestedText` **may be shorter** than `anchor`,
- `suggestedText` **may be longer** than `anchor`,
- `suggestedText` **may be empty** for deletion-like suggestions,
- but `suggestedText` must still represent the replacement for the `anchor` only.

It must **not** represent the fully rewritten `context` unless the `anchor` itself is that full context.

---

## Valid examples

### 1. Shorter replacement

```json
{
  "type": "track-change",
  "anchor": "completamente necesario",
  "context": "Era completamente necesario terminarlo hoy.",
  "suggestedText": "necesario",
  "justification": "‘Necesario’ ya implica completitud.",
  "category": "redundancia",
  "severity": "high"
}
```

Resulting text:

```text
Era necesario terminarlo hoy.
```

### 2. Longer replacement

```json
{
  "type": "track-change",
  "anchor": "llegó",
  "context": "Cuando llegó, todos guardaron silencio.",
  "suggestedText": "había llegado",
  "justification": "El pluscuamperfecto mejora la relación temporal con la acción posterior.",
  "category": "tiempo verbal",
  "severity": "medium"
}
```

Resulting text:

```text
Cuando había llegado, todos guardaron silencio.
```

### 3. Punctuation normalization with a shorter replacement

```json
{
  "type": "track-change",
  "anchor": "¡¿Ah?!",
  "context": "—¡¿Ah?! ¿Por qué habría de estar interesada en…?",
  "suggestedText": "¡Ah!",
  "justification": "Se recomienda evitar la combinación redundante de exclamación e interrogación cuando no aporta una intención distinta.",
  "category": "tipografía",
  "severity": "high"
}
```

Resulting text:

```text
—¡Ah! ¿Por qué habría de estar interesada en…?
```

### 4. Insertion-like replacement where the anchor stays small

```json
{
  "type": "track-change",
  "anchor": "dijo",
  "context": "Ella dijo que volvería antes del amanecer.",
  "suggestedText": "había dicho",
  "justification": "La frase requiere pluscuamperfecto por la secuencia temporal del relato.",
  "category": "tiempo verbal",
  "severity": "medium"
}
```

Resulting text:

```text
Ella había dicho que volvería antes del amanecer.
```

---

## Invalid example

The following payload is invalid for the current add-in contract:

```json
{
  "type": "track-change",
  "anchor": "¡¿Ah?!",
  "context": "—¡¿Ah?! ¿Por qué habría de estar interesada en…?",
  "suggestedText": "—¡Ah! ¿Por qué habría de estar interesada en…?",
  "justification": "...",
  "category": "tipografía",
  "severity": "high"
}
```

Why it is invalid:

- `anchor` points to only one fragment: `"¡¿Ah?!"`
- but `suggestedText` contains the whole rewritten sentence
- the frontend replace operation will replace only the anchor range with that full sentence
- the remaining text from the original context may remain around it, producing duplicated or malformed output

---

## Decision rule for the backend

When generating a `track-change` suggestion, the backend should follow this rule:

### Use `track-change` only when you can express the change as:

```text
replace(anchor) -> suggestedText
```

Where:

- `anchor` is an exact substring of `context`
- `suggestedText` is the exact replacement for that substring only

### Do not use `track-change` when the real edit is:

- a rewrite of the full sentence,
- a rewrite of multiple distant fragments at once,
- a transformation that cannot be expressed as one contiguous replacement.

In those cases, the backend should prefer one of these strategies:

1. emit a **smaller contiguous anchor** with a valid replacement, or
2. emit a **larger contiguous anchor** that covers the full span to rewrite, or
3. downgrade to **`comment-only`** if no safe contiguous replacement can be expressed.

---

## Safe generation checklist for backend

Before emitting a `track-change` suggestion, verify:

1. `context.includes(anchor)` is true,
2. `anchor` is the exact contiguous span to be replaced,
3. `suggestedText` is the replacement for that span only,
4. applying `context.replace(anchor, suggestedText)` would produce the intended local rewrite,
5. the suggestion does not depend on rewriting text outside the chosen anchor span.

If those conditions do not hold, the suggestion should not be emitted as a `track-change`.

---

## Summary

The key rule is:

> `anchor` defines the span to replace. `suggestedText` defines the replacement for that span.

`SuggestedText` can be longer, shorter, or empty.

What it cannot be is a full rewritten context when the `anchor` is only one small part of that context.
