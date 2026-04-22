# Tasks: Incremental Apply Snapshot

## Phase 1: Contractos y RED del batch real

- [ ] 1.1 Agregar contratos de snapshot, patch y rebase en `src/domain/types.ts`. Objetivo: modelar `AnalysisSnapshot`, `ApplyMutationPatch` y `RebasedSuggestion`.
- [ ] 1.2 Crear `src/adapters/word/BatchApplyOrchestrator.test.ts`. Objetivo: certificar con RED que `reverse()` no equivale a posición real de documento.
- [ ] 1.3 Extender `src/adapters/word/WordAdapterApplySuggestions.test.ts`. Objetivo: dejar claro que la suite actual de mock total es soporte, no escudo primario contra mislocation.

## Phase 2: Patch real desde Word

- [ ] 2.1 Modificar `src/adapters/word/ApplySuggestionCommand.ts`. Objetivo: devolver un patch localizado basado en la mutación real aplicada en Word.
- [ ] 2.2 Agregar tests de apply patch en `src/adapters/word/ApplySuggestionCommand*.test.ts`. Objetivo: cubrir delta de longitud y reread local del párrafo afectado.

## Phase 3: Snapshot viva y rebase incremental

- [ ] 3.1 Modificar `src/adapters/word/BatchApplyOrchestrator.ts`. Objetivo: construir snapshot inicial y actualizarla por versión tras cada apply exitoso.
- [ ] 3.2 Implementar rebase incremental de pendientes en el orquestador o colaborador extraído. Objetivo: desplazar offsets lejanos y marcar hot paragraphs para reread local.
- [ ] 3.3 Reemplazar la heurística `reverse()` por un seam explícito de ranking por posición snapshot/documento.

## Phase 4: Integración, fallback y validación

- [ ] 4.1 Ajustar `src/adapters/word/WordAdapter.ts` para proveer los insumos iniciales del snapshot sin releer todo el documento por cada apply.
- [ ] 4.2 Verificar que `src/adapters/word/ApplySuggestionCommandSearch.test.ts` siga verde como red de seguridad fallback.
- [ ] 4.3 Correr suites focalizadas de batch/apply/read y documentar hallazgos antes del commit.
