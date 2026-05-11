# Tasks: Semantic Reconciliation Telemetry

## Phase 1: Contratos y guardrails RED

- [ ] 1.1 Modificar `src/domain/types.ts` y `src/domain/ports.ts` para definir `ResolutionExecutionReport`, warnings de resolución y `ITelemetryPort` best-effort.
- [ ] 1.2 [TEST][REGRESSION] Agregar en `src/adapters/word/__tests__/WordAdapterAcceptSuggestion.test.ts` y `src/adapters/word/__tests__/WordAdapterRejectSuggestion.test.ts` casos RED de reconciliación semántica tras `ItemNotFound` tardío.
- [ ] 1.3 [TEST][REGRESSION] Agregar en `src/taskpane/__tests__/TaskpaneSuggestionResolution.test.ts` el guardrail de “terminal con warning” sin reactivar botones.

## Phase 2: Reconciliación semántica

- [ ] 2.1 Modificar `src/adapters/word/resolve-suggestion/ResolveSuggestionOperationalExecutor.ts` para devolver un reporte explícito de ejecución por intento.
- [ ] 2.2 Modificar `src/adapters/word/resolve-suggestion/ResolveSuggestionTrackChangeOrchestrator.ts` para re-observar la sugerencia y decidir `accepted|rejected|retryable` tras fallos.
- [ ] 2.3 Modificar `src/adapters/word/resolve-suggestion/ResolveSuggestionCommand.ts` para orquestar execute → reconcile → cleanup → inspect con prioridad a la verdad semántica.
- [ ] 2.4 Modificar `src/adapters/word/resolve-suggestion/SuggestionResolutionCleanup.ts` y `src/adapters/word/resolve-suggestion/ResolveSuggestionResultFactory.ts` para devolver warnings en vez de degradar éxito terminal a `error`.

## Phase 3: Workflow, UI y telemetría

- [ ] 3.1 Modificar `src/domain/suggestion/SuggestionResolutionWorkflow.ts` y `src/domain/suggestion/SuggestionStateMachine.ts` para propagar éxito terminal con warnings y mantener feedback basado en outcome semántico.
- [ ] 3.2 Modificar `src/taskpane/SuggestionCardRenderer.ts` para renderizar warnings terminales sin dejar la card retryable.
- [ ] 3.3 Crear `src/adapters/telemetry/ConsoleTelemetryAdapter.ts` y cablearlo desde `src/taskpane/taskpane.ts` como sink por defecto no bloqueante.
- [ ] 3.4 Emitir eventos estructurados desde `ResolveSuggestionCommand` y colaboradores para `observe-before`, `execute`, `reconcile`, `cleanup` e `inspect-after`.

## Phase 4: Documentación y validación

- [ ] 4.1 Crear `docs/TELEMETRY.md` con principios, envelope de eventos, correlación y reglas de no interferencia.
- [ ] 4.2 Actualizar `docs/review-domain-and-track-changes.md` o `docs/architecture.md` si hace falta referenciar la nueva regla de reconciliación semántica.
- [ ] 4.3 Ejecutar suites focalizadas de adapter/taskpane para accept, reject y warnings terminales antes de cualquier commit.
