# Postmortem — Replace Resolution Atomicity Bug

## Estado

- **Bug NO resuelto todavía en Word real**.
- El objetivo sigue siendo el mismo desde el primer reporte: si un replace no puede resolverse atómicamente, **no** debe dejar Word en estado intermedio ni marcar la tarjeta como `accepted` / `rejected`.
- Este documento registra **todos los intentos de fix** hasta ahora, incluso los fallidos o parciales, para poder borrarlos después uno por uno hasta dejar solo el fix real.

---

## Resumen ejecutivo

El bug empezó como un problema de **falso success**: el add-in podía certificar `accepted` / `rejected` aunque solo una mitad semántica del replace hubiera sido resuelta en Word.

Después, al endurecer la observación, el problema cambió a **falso error**: el add-in ya no mentía con success, pero tampoco lograba resolver replaces que Word sí seguía mostrando como pendientes.

Con cada iteración se fue aislando una capa distinta del problema:

1. observación incompleta,
2. uso incompleto de metadata `compound-v2`,
3. orden semántico de ejecución,
4. falta de `sync` entre pasos,
5. reuso de proxies stale,
6. sobre-observación de tracked changes,
7. duplicación semántica de lados,
8. re-observación demasiado permisiva,
9. verificación post-ejecución sin rollback real.

La conclusión actual es dura pero clara:

> El flujo de resolución de replace **no es transaccional**. Puede detectar que la atomicidad se perdió, pero lo hace **después** de haber mutado Word. Y hoy no existe un mecanismo real de rollback/snapshot para devolver el documento al estado anterior.

---

## Requisito de producto / contrato

El contrato esperado para `accept` / `reject` de una sugerencia replace es:

- si Word resuelve el replace completo, el documento queda correcto y la tarjeta pasa a estado terminal;
- si Word no lo resuelve completo, el documento **no** debe quedar medio mutado;
- el add-in **no** debe limpiar comment/CC ni enviar feedback terminal si el replace quedó parcial.

---

## Cronología de intentos de fix

### Intento 1 — Exigir par completo `Added + Deleted` antes de certificar replace

**Idea**

Evitar falso success endureciendo `SuggestionResolutionObserver`: un replace solo se considera `confirmed-pending` si Word expone **ambos** lados semánticos.

**Cambios principales**

- `SuggestionResolutionObserver.ts`
  - helper `hasCompleteReplaceTrackedChangePair(...)`
  - replace inicial: `trackedChanges.length > 0` → requiere `Added + Deleted`
- tests nuevos en accept/reject para el caso medio visible (`s-half-visible`) → debe devolver `unobservable`

**Qué arregló**

- evitó que el add-in dijera `accepted` / `rejected` cuando solo se veía un lado del replace.

**Qué no arregló**

- Word real seguía mostrando replaces pendientes que el add-in no lograba resolver.

**Diagnóstico**

- fue un fix **correcto para el falso success inicial**, pero insuficiente para resolver el bug real.

---

### Intento 2 — Permitir recovery después de progreso parcial

**Idea**

Si un lado ya se ejecutó, una re-observación posterior puede mostrar solo el lado restante. El recovery no debía exigir de nuevo un par completo.

**Cambios principales**

- `ResolveSuggestionCommand.ts`
  - recovery dejó de requerir `observationStatus === confirmed-pending`
  - pasó a aceptar cualquier re-observación con tracked changes restantes y sin `identity-lost`

**Qué arregló**

- evitó bloquear retries válidos después de ejecutar una mitad.

**Qué no arregló**

- no resolvió el problema de fondo en Word real.

**Diagnóstico**

- fix parcial, preservado como necesario pero no suficiente.

---

### Intento 3 — Corregir fixtures que accidentalmente parecían replace `compound-v2`

**Idea**

Varios tests legacy se rompieron no por el fix, sino porque sus fixtures seguían pareciendo replace aunque en realidad representaban casos simples/no-replace.

**Cambios principales**

- ajustes de `ccTitle`, `anchor`, `suggestedText`, `context` en tests de accept/reject
- evitar que helpers auto-generen `compound-v2` en escenarios que no lo merecen

**Qué arregló**

- limpió ruido de suite.

**Qué no arregló**

- no toca el bug real; solo mejoró la fidelidad de tests.

---

### Intento 4 — Usar explícitamente `deletedSideRef` como locator de resolución

**Idea**

`compound-v2` no debía ser solo metadata decorativa. Si el lado borrado no aparece por CC/body proximity, usar `deletedSideRef` para re-localizarlo explícitamente.

**Cambios principales**

- `ReplaceIdentityParser.ts`
  - `getDeletedSideLocator(...)`
  - `getOperationalAnchorLocator(...)`
- `SuggestionResolutionObserver.ts`
  - búsqueda explícita del deleted side en `context.document.body`
  - merge de tracked changes encontrados por deleted-side locator
- tests accept/reject nuevos para el caso donde el `Deleted` solo aparece por ese locator

**Qué arregló**

- mejoró la observación de replaces que Word expone de forma rara.

**Qué no arregló**

- Word real seguía fallando durante la resolución.

**Diagnóstico**

- el metadata `compound-v2` ayuda a **re-localizar**, pero no alcanza para garantizar ejecución correcta.

---

### Intento 5 — Imponer orden semántico de ejecución

**Idea**

No confiar en el orden arbitrario del host. Un replace debe ejecutarse por semántica, no por enumeración:

- `accept`: `Deleted` → `Added`
- `reject`: `Added` → `Deleted`

**Cambios principales**

- `TrackedChangeResolutionExecutor.ts`
  - `orderTrackedChangesForExecution(...)`
- tests de orden explícito en accept/reject

**Qué arregló**

- dejó de depender del orden de Word.

**Qué no arregló**

- el bug real persistió.

**Diagnóstico**

- el orden era importante, pero tampoco era la causa final.

---

### Intento 6 — Hacer `sync` después de cada paso semántico

**Idea**

No acumular ambos `accept()` / `reject()` antes del siguiente `context.sync()`. Forzar commit host-step-by-host-step.

**Cambios principales**

- `TrackedChangeResolutionExecutor.apply()` pasó a async y recibe `context`
- `await context.sync()` después de cada tracked change
- tests nuevos de secuencia `step -> sync -> step -> sync`

**Qué arregló**

- hizo el flujo más realista frente al host.

**Qué no arregló**

- el segundo paso seguía cayendo con `ItemNotFound` en Word.

**Diagnóstico**

- `sync` por paso era necesario, pero no suficiente; el segundo proxy seguía stale.

---

### Intento 7 — Re-observar antes del segundo lado con proxies frescos

**Idea**

Después del primer paso y `sync`, no reutilizar el proxy inicial del segundo lado. Re-localizar y re-observar desde Word.

**Cambios principales**

- `ResolveSuggestionCommand.ts`
  - `executeReplaceTrackedChangesWithReobservation(...)`
  - `executeReplaceSemanticStep(...)`
  - `reobserveResolutionCandidates(...)`
- tests host-realistic donde el segundo lado solo existe como proxy fresco después del primer `sync`

**Qué arregló**

- acercó el modelo a lo que hace Word de verdad.

**Qué no arregló**

- el segundo paso todavía podía lanzar `ItemNotFound` incluso después de re-observar.

**Diagnóstico**

- el stale-proxy problem existía, pero había más de una forma de fallar.

---

### Intento 8 — Interpretar `ItemNotFound` post-segundo-paso como posible éxito semántico

**Idea**

Si el host ya aplicó el cambio y después invalida el proxy, `ItemNotFound` no siempre significa fracaso semántico. Re-observar tras el error y, si el lado ya no está, tratarlo como éxito.

**Cambios principales**

- refinamiento en `executeReplaceSemanticStep(...)`
  - recovery observation tras error
  - final recovery observation tras retry
  - si el lado ya no está, marcar `completed: true`

**Qué arregló**

- redujo falsos negativos basados en proxies inválidos.

**Qué no arregló**

- siguió sin resolver el bug real en Word.

**Diagnóstico**

- otro fix parcial válido, pero no final.

---

### Intento 9 — Logging estructurado de locator/observer/executor/cleanup

**Idea**

Antes de seguir parcheando a ciegas, registrar exactamente qué candidato se selecciona, qué fuentes observaron tracked changes, qué step ejecutó el executor y en qué fase explota.

**Cambios principales**

- `SuggestionLocator.ts`
  - candidate diagnostics estructurados
  - warning para duplicates indistinguibles
- `ResolutionContext.ts`
  - `ResolutionObservationDebugMetadata`
- `SuggestionResolutionObserver.ts`
  - metadata por fuente: cc, ccRange, body, deletedSide, anchor, comment
- `ResolveSuggestionCommand.ts`
  - logs ricos en locate/observe/execute/cleanup/catch
- `TrackedChangeResolutionExecutor.ts`
  - logs por step / queue / sync
- `SuggestionResolutionCleanup.ts`
  - logs para comment/anchor cleanup

**Qué arregló**

- no arregló el bug, pero permitió aislarlo mejor.

**Hallazgo clave**

- hubo sesiones donde el locator sí elegía el candidate correcto y el problema estaba en observación/ejecución posterior.

---

### Intento 10 — Reducir sobre-observación de tracked changes

**Idea**

No mezclar todo lo que aparece en cc, ccRange, body, comment, anchor, deletedSide. Elegir la combinación mínima que ya forme un replace válido.

**Cambios principales**

- `SuggestionResolutionObserver.ts`
  - buckets separados por fuente
  - `resolveTrackedChangesForReplace(...)`
  - prioridad por combinaciones mínimas útiles
- tests accept/reject nuevos para evitar over-collection con evidencias stale

**Qué arregló**

- bajó ejecuciones infladas (ej. 8 pasos para un replace simple).

**Qué no arregló**

- seguían apareciendo duplicados semánticos del mismo lado.

---

### Intento 11 — Deduplicación semántica por lado (`Added` / `Deleted`)

**Idea**

No deduplicar por proxy/id, sino por **slot semántico**. Para un replace simple debe haber como máximo:

- un `Deleted`
- un `Added`

**Cambios principales**

- `SuggestionResolutionObserver.ts`
  - `normalizeReplaceTrackedChanges(...)`
  - combinación normalizada antes de evaluar completitud
- tests accept/reject nuevos para duplicate semantic sides

**Qué arregló**

- dejó el count correcto en 2 pasos semánticos.

**Qué no arregló**

- el segundo paso seguía fallando.

---

### Intento 12 — Re-observación side-specific del lado restante

**Idea**

Después del primer paso, no volver a exigir un par completo; observar solo el lado semántico restante.

**Cambios principales**

- `SuggestionResolutionObserver.ts`
  - `observeReplaceSuggestionSemanticSide(...)`
  - `observeResolutionCandidatesForSemanticSide(...)`
  - prioridades distintas para `Added` y `Deleted`
- `ResolveSuggestionCommand.ts`
  - re-observación side-specific antes del segundo paso y en recovery
- tests host-realistic nuevos para:
  - remaining `Added` sin requerir `Deleted`
  - remaining `Deleted` sin requerir `Added`

**Qué arregló**

- permitió completar algunos escenarios que antes requerían el par completo injustificadamente.

**Qué rompió**

- en Word real apareció **falso success**: una mitad quedaba pendiente, pero el flujo seguía con cleanup, inspect-after, feedback y tarjeta terminal.

**Diagnóstico**

- este intento fue especialmente peligroso porque rompió atomicidad por permisivo.

---

### Intento 13 — Verificación post-ejecución antes de cleanup

**Idea**

Antes de borrar comment/CC y marcar success, re-observar el replace completo. Si Word todavía muestra tracked changes pendientes, bloquear terminal success.

**Cambios principales**

- `ResolveSuggestionCommand.ts`
  - `verifyReplaceResolvedBeforeCleanup(context)`
  - bloquea cleanup/accepted/rejected si sigue habiendo tracked changes o `identity-lost`
  - metadata nueva en telemetry de execute
- tests nuevos:
  - accept: no cleanup si queda una mitad pendiente
  - reject: no cleanup si queda una mitad pendiente

**Qué arregló**

- conceptualmente ataca el falso success correcto.

**Qué rompió**

- en tests re-observaba fixtures estáticos y empezaba a clasificar casi todo como parcial.

**Diagnóstico**

- la idea sigue siendo correcta, pero la infraestructura de mocks no estaba preparada.

---

### Intento 14 — Hacer el helper de Office.js más host-realistic para verificación post-ejecución

**Idea**

Si un lado ya fue aceptado/rechazado en el mock, posteriores observaciones no deben seguir viéndolo como pendiente salvo que el test lo reintroduzca explícitamente.

**Cambios principales**

- `WordAdapterActionTestHelper.ts`
  - arrays mutables por fuente
  - `trackedChangeCollections`
  - `removeTrackedChangesBySemanticSide(...)`
  - wrapping de `accept()` / `reject()` para simular progreso host
  - luego refinamiento para también incluir colecciones per-CC

**Qué arregló**

- estabilizó bastante la suite para la nueva guardia de atomicidad.

**Qué no arregló**

- aún así quedaron problemas con continuidad de candidate lógico tras re-observación en escenarios de duplicate-tag.

---

### Intento 15 — Intento de preservar continuidad del candidate activo en re-observación

**Idea**

No confiar ciegamente en `relocated.selectedCc`; intentar seguir el mismo candidate usado por la observación activa durante partial execution / recovery.

**Cambios principales**

- `ResolveSuggestionCommand.ts`
  - `reobserveResolutionCandidates(...)` y `reobserveResolutionCandidatesForSemanticSide(...)` aceptan `preferredCc`
  - reordenan candidatos poniendo primero el candidate preferido
  - call sites actualizados para pasar `observation.selectedCc` / `activeObservation.selectedCc`

**Qué arregló**

- redujo parte del drift entre observación activa y re-locación.

**Qué no arregló**

- seguían fallando tests de duplicate-candidate / stale-candidate / later-candidate.

**Diagnóstico**

- el error estaba en usar **identidad de objeto** para un problema que requiere **identidad lógica persistida**.
- los candidatos re-localizados son proxies frescos; comparar por referencia no alcanza.

---

## Qué sabemos hoy con bastante certeza

### 1. No es solo un bug

Es una cadena de problemas superpuestos:

- observación,
- continuidad de candidate,
- stale proxies,
- verificación tardía,
- ausencia de rollback.

### 2. `compound-v2` ayuda, pero no resuelve atomicidad

Sirve como memoria de identidad:

- `insertedSideRef`
- `deletedSideRef`
- `anchorRef`

Pero no provee:

- continuidad segura de proxy,
- snapshot previo,
- rollback.

### 3. Accept y reject comparten esqueleto, pero no son simétricos

- orden semántico distinto,
- prioridades de observación side-specific distintas,
- puntos de falla distintos.

### 4. El sistema actual NO implementa rollback real

No hay:

- Memento,
- snapshot/restore,
- undo propio,
- compensating restore robusto.

Entonces cuando el flujo detecta partial state, muchas veces ya es tarde: Word ya quedó medio mutado.

---

## Causa técnica actual del bug

La causa de fondo, hasta donde llega la evidencia hoy, es esta:

> El replace se resuelve como una secuencia de mutaciones host no transaccionales. El sistema intenta reconstruir atomicidad con observación + re-observación + verificación posterior, pero no posee continuidad lógica suficientemente fuerte del artifact ni rollback real cuando Word ya persistió una mitad.

En castellano brutal:

- tocamos Word paso a paso,
- Word puede re-materializar/invalidate proxies,
- nosotros volvemos a observar con heurísticas,
- a veces encontramos algo compatible pero no exactamente el artifact correcto,
- y aunque detectemos el desastre, no tenemos cómo deshacer lo ya synced.

---

## Impacto

- Word puede quedar en **estado parcial**.
- El taskpane puede caer en dos errores distintos para el mismo bug:
  - partial state detectado tarde,
  - remaining side no reexpuesto.
- En intentos permisivos, incluso se llegó a `accepted` / cleanup / feedback con Word todavía medio pendiente.
- El tiempo de diagnóstico se disparó porque cada capa arreglada reveló otra capa más abajo.

---

## Qué NO hay que olvidar

1. Varias de estas correcciones parciales fueron **válidas y necesarias**.
   - No fueron “inútiles”; simplemente no eran el fix final.

2. Los intentos fallidos son evidencia forense.
   - no borrarlos todavía,
   - documentarlos primero,
   - luego eliminar solo cuando el fix real esté probado en Word.

3. El problema no es solo “falta de memoria”.
   - hay memoria de identidad,
   - falta memoria **transaccional** y memoria **histórica operativa**.

---

## Estado actual

- Hay múltiples fixes parciales en el código que siguen en evaluación.
- La línea de trabajo más sana hoy es:
  - mantener la guardia contra falso success,
  - preservar candidate continuity por identidad lógica, no por referencia de proxy,
  - y no seguir confundiendo detección tardía con rollback real.

**Pero la verdad dura es esta:**

> Mientras no exista una estrategia real de snapshot/restore o rollback compensatorio confiable, la atomicidad fuerte en Word seguirá siendo estructuralmente frágil.

---

## Archivos más relevantes en esta investigación

- `src/adapters/word/ResolveSuggestionCommand.ts`
- `src/adapters/word/resolution/SuggestionResolutionObserver.ts`
- `src/adapters/word/resolution/TrackedChangeResolutionExecutor.ts`
- `src/adapters/word/resolution/SuggestionLocator.ts`
- `src/adapters/word/WordAdapterActionTestHelper.ts`
- `src/adapters/word/WordAdapterAcceptSuggestion.test.ts`
- `src/adapters/word/WordAdapterRejectSuggestion.test.ts`
- `src/adapters/word/ApplySuggestionCommand.ts`
- `src/adapters/word/ReplaceIdentityParser.ts`

---

## Nota final

Este documento NO certifica un fix. Certifica algo más importante en este momento:

**qué probamos, qué rompió, qué mejoró, y por qué todavía no alcanza.**
