# CLAUDE.md — Word Track Changes Add-in

Documento de planificación para Claude Code. Leer completo antes de escribir cualquier código.

---

## Tracer Bullets

Cuando construyas features, construye primero una rebanada pequeña y end-to-end de la funcionalidad, busca feedback, y luego expande desde ahí.

Las tracer bullets vienen de The Pragmatic Programmer. Cuando construyes sistemas, quieres código que te dé feedback lo más rápido posible. Son rebanadas pequeñas de funcionalidad que atraviesan todas las capas del sistema a la vez, permitiéndote testear y validar tu enfoque temprano. Esto ayuda a identificar problemas potenciales y asegura que la arquitectura general es sólida antes de invertir tiempo significativo en desarrollo.

**No construyas capas horizontales en aislamiento. No construyas toda la UI, todos los modelos, todo el manejo de errores antes de validar que la conexión central funciona. Construye una rebanada vertical pequeña, testéala en Word real, obtén feedback, y muévete a la siguiente en un contexto fresco.**

---

## 1. Visión general del proyecto

Add-in de Microsoft Word que analiza un documento y propone sugerencias editoriales mediante el sistema nativo de **Track Changes (Control de Cambios)**. El usuario acepta o rechaza cada sugerencia directamente desde la pestaña "Revisar" de Word, sin aprender ninguna interfaz nueva.

El problema de autoría (las revisiones aparecen con el nombre del usuario, no del add-in) está aceptado y es un no-requisito. No se intentará solucionar con OOXML.

---

## 2. Stack tecnológico

Base: plantilla oficial `Office-Addin-TaskPane` con TypeScript + Webpack.

```
office-addin-taskpane (package.json base)
├── TypeScript 5.4
├── Webpack 5 (bundler)
├── office-addin-debugging (sideload en Word desktop)
├── @types/office-js (tipado de la API)
└── Sin React — UI en HTML/CSS vanilla en el Task Pane
```

Sin React: el `package.json` base no lo incluye y la UI del Task Pane es mínima. Añadir React sería sobreingeniería. Si la UI crece, agregar React es trivial sobre este stack.

---

## 3. Requisitos de API (Requirement Sets)

| Feature necesaria | Requirement Set mínimo |
|---|---|
| Activar/leer `changeTrackingMode` | **WordApi 1.4** |
| `getTrackedChanges()` en Body/Range | **WordApi 1.6** |
| `TrackedChange.accept()` / `reject()` | **WordApi 1.6** |
| `insertText()` con tracking activo | **WordApi 1.1** |

El add-in requiere **WordApi 1.6** como mínimo. Word Online tiene soporte completo sin restricciones de versión, lo que cubre usuarios con clientes de escritorio más antiguos.

---

## 4. Arquitectura del flujo principal

```
[Usuario abre Task Pane]
        │
        ▼
[Add-in lee el documento completo]
        │
        ▼
[Lógica de análisis produce lista de sugerencias]
  { originalText, suggestedText, justification }
        │
        ▼
[Add-in preserva changeTrackingMode actual]
[Activa changeTrackingMode = "TrackAll"]
        │
        ▼
[Por cada sugerencia: localiza Range → insertText(replace)]
  Word registra el cambio como revisión automáticamente
        │
        ▼
[Add-in restaura changeTrackingMode al estado previo]
        │
        ▼
[Usuario revisa con UI nativa de Word (pestaña Revisar)]
```

### Patrón crítico: preserve-and-restore

Nunca asumir que `changeTrackingMode` estaba en `"Off"`. Siempre leer, guardar, y restaurar.

```typescript
await Word.run(async (context) => {
  context.document.load("changeTrackingMode");
  await context.sync();
  const previousMode = context.document.changeTrackingMode;

  context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  await context.sync();

  // ... insertar sugerencias ...

  context.document.changeTrackingMode = previousMode;
  await context.sync();
});
```

---

## 5. Estructura de archivos

```
/
├── manifest.xml
├── webpack.config.js
├── tsconfig.json
├── package.json
│
└── src/
    ├── taskpane/
    │   ├── taskpane.html
    │   ├── taskpane.css
    │   └── taskpane.ts
    │
    └── lib/
        ├── wordApi.ts      # Toda interacción con Office.js. Nunca contiene lógica de negocio.
        ├── analyzer.ts     # Lógica de análisis. Nunca importa Office ni Word.
        └── types.ts        # Interfaces TypeScript compartidas
```

---

## 6. Interfaces TypeScript (`types.ts`)

```typescript
export interface Suggestion {
  id: string;
  originalText: string;
  suggestedText: string;
  justification: string;
  paragraphIndex?: number;  // Para filtrar si hay múltiples ocurrencias
}

export interface InsertionResult {
  successCount: number;
  failedSuggestions: Suggestion[];
}
```

---

## 7. Funciones clave de `wordApi.ts`

```typescript
// Lee el texto completo del documento
export async function getDocumentText(): Promise<string>

// Inserta sugerencias como revisiones trazables.
// IMPORTANTE: encolar todas las inserciones en un solo Word.run,
// no hacer un context.sync() por cada sugerencia.
export async function insertSuggestionsAsTrackedChanges(
  suggestions: Suggestion[]
): Promise<InsertionResult>
```

---

## 8. Manifest (`manifest.xml`) — puntos críticos

```xml
<Requirements>
  <Sets DefaultMinVersion="1.6">
    <Set Name="WordApi" MinVersion="1.6" />
  </Sets>
</Requirements>

<Permissions>ReadWriteDocument</Permissions>

<Hosts>
  <Host Name="Document" />
</Hosts>
```

Usar formato XML v1.1 (no el formato unificado JSON de Teams).

---

## 9. Gotchas conocidos

**Bug: `TrackedChange.text` vacío en Mac para tipo "Deleted"** (issue #5188, dic 2024).
Solo afecta lectura de cambios existentes. No impacta el flujo principal de inserción.
Mitigación futura: `if (change.type === "Deleted" && change.text === "") { /* usar getRange().text */ }`

**Bug: `getTrackedChanges()` falla si el documento tiene cambios tipo "Move"** (issue #5535, mar 2025).
Envolver toda llamada a `getTrackedChanges()` en try/catch. El flujo de inserción no se ve afectado.

**Búsqueda con múltiples ocurrencias**: `body.search()` retorna todas. Tomar `items[0]`. Si se requiere precisión de párrafo, usar `paragraphIndex` de `Suggestion`.

**Performance**: más de ~50 sugerencias puede ser lento. Mostrar spinner y procesar en lotes de 10-15 si es necesario.

---

## 10. Comandos de desarrollo

```bash
npm install       # Instalar dependencias
npm start         # Dev server + sideload en Word desktop
npm stop          # Detener
npm run build     # Build de producción
npm run validate  # Validar manifest
```

---

## 11. Plan de implementación por tracer bullets

Cada bullet es una rebanada vertical completa y testeable en Word real. No avanzar al siguiente sin validar el actual. Cada bullet se trabaja en un contexto fresco.

---

### Bullet 1 — El add-in carga y lee el documento

**Objetivo**: confirmar que el sideload funciona y que Office.js puede comunicarse con el documento.

Construir:
- `manifest.xml` con IDs, permisos (`ReadWriteDocument`) y requirement set 1.6
- `types.ts` con las interfaces base
- `wordApi.ts` con solo `getDocumentText()`
- `taskpane.html` con un solo botón: "Leer documento"
- `taskpane.ts` que al hacer clic muestra los primeros 200 caracteres en un `<pre>`

**Criterio de éxito**: el Task Pane muestra texto real del documento. Si esto no funciona, nada más funciona.

---

### Bullet 2 — Una sugerencia hardcodeada aparece como revisión

**Objetivo**: confirmar el ciclo completo tracking → inserción → revisión visible.

Construir:
- `wordApi.ts`: agregar `insertSuggestionsAsTrackedChanges()`
- `taskpane.html`: agregar botón "Sugerencia de prueba"
- `taskpane.ts`: al hacer clic, llama a la función con una sugerencia hardcodeada (palabra concreta del documento de prueba)

**Criterio de éxito**: la palabra aparece tachada y el reemplazo subrayado en Word. El usuario puede aceptarla o rechazarla desde la pestaña "Revisar". La pestaña "Revisar" muestra la revisión correctamente.

---

### Bullet 3 — El analizador produce sugerencias reales

**Objetivo**: reemplazar el hardcode con lógica real.

Construir:
- `analyzer.ts` con la lógica real de análisis (o llamada a API externa)
- Conectar: `getDocumentText()` → `analyzer.ts` → `insertSuggestionsAsTrackedChanges()`
- UI: botón único "Analizar y sugerir" con estado de carga

**Criterio de éxito**: el add-in analiza un documento real y las sugerencias del analizador aparecen como revisiones en Word.

---

### Bullet 4 — Reporte de resultados y manejo de errores

**Objetivo**: que el add-in informe qué funcionó y no se rompa ante condiciones adversas.

Construir:
- Mostrar en Task Pane: cuántas sugerencias se insertaron, cuáles no se encontraron
- try/catch en todas las llamadas a `Word.run`
- Mensaje útil si el documento está vacío o protegido

**Criterio de éxito**: probar con un documento donde una sugerencia no existe en el texto. El add-in reporta el fallo sin romperse.

---

### Bullet 5 — Validación en Word Online

**Objetivo**: confirmar consistencia fuera del cliente de escritorio.

Construir: nada. Solo testing.

**Criterio de éxito**: bullets 1–4 pasan en Word Online sin modificar el código.

---

## 12. Límites de alcance (decisiones tomadas)

- No modifica el autor de las revisiones — no-requisito aceptado
- No acepta/rechaza revisiones programáticamente — delega a la UI nativa de Word
- No usa inyección OOXML — innecesario dado lo anterior
- No soporta documentos con protección DRM/RMS — limitación conocida, documentar al usuario
- No soporta coautoría activa durante el análisis — limitación conocida
