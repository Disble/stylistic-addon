# Backlog

## Arquitectura futura

- **Memoria útil de largo plazo para el add-in**
  - Problema observado: el add-in persiste identidad mínima en Word (`compound-v2`, tags, comments), pero no conserva suficiente memoria operativa/transaccional para reconocer historia de artifacts, validar evolución del host ni revertir cambios parciales.
  - Impacto: complica resolución atómica, diagnóstico de artifacts stale y rollback real cuando Word queda medio mutado.
  - Dirección deseada: diseñar memoria de sistema más fuerte (historial de runtime + snapshot/restore transaccional) en lugar de depender solo de re-observación heurística.
  - Prioridad actual: **no tocar ahora**; primero matar el bug de atomicidad/rollback en Word.
