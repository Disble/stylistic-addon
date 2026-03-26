# Propuesta: Agregar `justification` al FeedbackPayload

## Contexto

Durante las pruebas alpha del `feedback-workflow` se detectó un edge case de clasificación incorrecta. El agente de feedback clasificó como `INTOCABLE` un comentario que debería haber sido ignorado, por falta de contexto sobre la naturaleza de la sugerencia original.

### Caso concreto

**Payload recibido:**
```json
{
  "category": "ortografia",
  "originalText": "¿Viento? acasp tiene",
  "suggestedText": "¿Viento? Acaso tiene",
  "rating": "negative",
  "severity": "high",
  "comment": "Es parte de un dialogo donde esta tartamudeando el personaje, se supone que pronuncia mal las palabras"
}
```

**Comportamiento actual:** El agente clasificó el comentario como `INTOCABLE` y escribió en el perfil del autor que los errores ortográficos en diálogos son intencionales — creando un precedente peligroso que podría blindar futuros errores reales.

**Causa raíz:** El agente no tenía acceso a la `justification` que generó el corrector (`stylisticAgent`). Con ese dato — por ejemplo: *"Error ortográfico normativo: 'acasp' es una errata de 'acaso'"* — el agente habría podido razonar correctamente: el corrector marcó esto como error normativo, el autor explica que es intencional por el personaje → el comentario es contextual a esa escena, no un rasgo de estilo del autor → no actualizar el perfil.

---

## Propuesta

Agregar el campo `justification` al `FeedbackPayload` en el frontend.

### Schema actual

```typescript
interface FeedbackPayload {
  category: string;
  originalText: string;
  suggestedText: string;
  rating: "positive" | "negative";
  severity: "high" | "medium" | "low";
  comment?: string;
}
```

### Schema propuesto

```typescript
interface FeedbackPayload {
  category: string;
  originalText: string;
  suggestedText: string;
  rating: "positive" | "negative";
  severity: "high" | "medium" | "low";
  justification: string;   // ← NUEVO: justificación del corrector
  comment?: string;
}
```

### Origen del campo

`justification` ya existe en cada `Suggestion` devuelta por el `stylistic-workflow`:

```typescript
// Suggestion (ya disponible en el frontend)
{
  originalText: string;
  suggestedText: string;
  justification: string;   // ← este campo
  category: string;
  severity: "high" | "medium" | "low";
}
```

El frontend ya tiene acceso a `suggestion.justification` en el momento en que el usuario hace ✅/❌. Solo hay que incluirlo en el payload que se envía al `FeedbackAdapter`.

### Cambio en el frontend

Un solo campo adicional al construir el payload en `taskpane.ts`:

```typescript
const payload: FeedbackPayload = {
  category: suggestion.category,
  originalText: suggestion.originalText,
  suggestedText: suggestion.suggestedText,
  justification: suggestion.justification,   // ← agregar esta línea
  rating: "positive",
  severity: suggestion.severity,
  ...(comment ? { comment } : {}),
};
```

---

## Impacto en el backend

El `feedback-workflow` deberá:

1. Agregar `justification: z.string()` al `feedbackWorkflowInputSchema` (Zod)
2. Incluir `justification` en el prompt que recibe el `feedbackAgent`

Con este dato, el agente puede razonar:

> *"El corrector justificó esta sugerencia como un error ortográfico normativo. El autor dice que es intencional por el personaje. Conclusión: el autor está explicando el contexto narrativo de esta instancia específica, no declarando una preferencia de estilo propia → no actualizar el perfil."*

---

## Prioridad

**Alta** — sin este campo, el agente de feedback opera con información incompleta y puede generar entradas incorrectas en el perfil del autor, degradando la calidad del sistema en lugar de mejorarla.

---

## Coordinación requerida

- **Frontend**: agregar `justification` al `FeedbackPayload` y al `FeedbackAdapter`
- **Backend**: actualizar `feedbackWorkflowInputSchema` y el prompt del `feedbackAgent`
- **Contrato**: actualizar `docs/api-contract.md` sección "Feedback Workflow Contract" con el nuevo campo
