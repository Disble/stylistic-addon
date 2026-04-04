# Tasks: Replace Suggestion Identity

## Phase A: Stop false `already-resolved` first

- [ ] 1.1 Add non-terminal resolution statuses to the domain contract. Objetivo: introducir `unobservable` y reservar `already-resolved` para confirmación positiva. Toca: `src/domain/types.ts`, `src/domain/ports.ts`.
- [ ] 1.2 Change Word observation fallback for legacy replace suggestions. Objetivo: mapear `0 tracked changes observed` a `unobservable` sin borrar anchors/comentarios. Toca: `src/adapters/word/WordAdapter.ts`, `src/adapters/word/cleanup/CommentCleanup.ts`.
- [ ] 1.3 Update resolution workflow semantics. Objetivo: omitir feedback y devolver resultado retryable cuando la observación sea ambigua. Toca: `src/domain/suggestion/SuggestionResolutionWorkflow.ts`, `src/domain/review/ReviewSessionMediator.ts`.
- [ ] 1.4 Extend review state handling in domain/UI. Objetivo: tratar `unobservable` como no terminal y mostrar estado explícito en taskpane. Toca: `src/domain/suggestion/SuggestionStateMachine.ts`, `src/taskpane/taskpane.ts`.
- [ ] 1.5 [TEST][REGRESSION] Add Tier 1 guard for the lie stop. Objetivo: cubrir `0 observed -> unobservable` y confirmar que no se reporta `already-resolved`. Toca: `src/adapters/word/WordAdapterAcceptSuggestion.test.ts`, `src/adapters/word/WordAdapterRejectSuggestion.test.ts`.
- [ ] 1.6 [TEST][REGRESSION] Add workflow/UI guards for ambiguous states. Objetivo: verificar feedback skipped, transición retryable y render taskpane para `unobservable`. Toca: `src/domain/suggestion/SuggestionResolutionWorkflow.test.ts`, `src/domain/suggestion/SuggestionStateMachine.test.ts`, `src/taskpane/TaskpaneSuggestionResolution.test.ts`.

## Phase B: Redesign replace identity as compound identity

- [ ] 2.1 Add compound replace identity types. Objetivo: modelar `SuggestionObservationStatus`, `WordArtifactRef` y `ReplaceSuggestionIdentity` versionada. Toca: `src/domain/types.ts`, comentarios de contrato en `src/domain/ports.ts`.
- [ ] 2.2 Write v2 replace metadata on apply. Objetivo: hacer que nuevas replace suggestions persistan refs `inserted-side`, `deleted-side` y `operational-anchor` sin asumir `CC == identity`. Toca: `src/adapters/word/ApplySuggestionCommand.ts`, helpers/tests asociados.
- [ ] 2.3 Implement dual-read observation in WordAdapter. Objetivo: leer `legacy-v1` y `compound-v2`, clasificar `confirmed-pending|confirmed-resolved|unobservable|identity-lost`, y mutar solo con observación suficiente. Toca: `src/adapters/word/WordAdapter.ts`.
- [ ] 2.4 Wire backward compatibility through workflow/state/UI. Objetivo: mantener artifacts viejos accionables, degradar legacy sin evidencia a `unobservable`, y exponer `identity-lost` como warning explícito. Toca: `src/domain/suggestion/SuggestionResolutionWorkflow.ts`, `src/domain/suggestion/SuggestionStateMachine.ts`, `src/domain/review/ReviewSessionMediator.ts`, `src/taskpane/taskpane.ts`.
- [ ] 2.5 [TEST][REGRESSION] Add Tier 1 compatibility matrix. Objetivo: cubrir bare ID, `stylistic:{type}:{id}`, metadata `compound-v2`, y refs faltantes/corruptas. Toca: `src/adapters/word/WordAdapterActionTestHelper.ts`, `src/adapters/word/ApplySuggestionCommand*.test.ts`, `src/adapters/word/WordAdapter*.test.ts`.
- [ ] 2.6 [TEST][REGRESSION] Add domain/taskpane guards for new semantics. Objetivo: verificar `identity-lost`, suppress feedback, y estados visibles no terminales/terminal-warning. Toca: `src/domain/suggestion/SuggestionResolutionWorkflow.test.ts`, `src/domain/suggestion/SuggestionStateMachine.test.ts`, `src/taskpane/TaskpaneSuggestionResolution.test.ts`, `src/taskpane/TaskpaneFeedback.test.ts`.
- [ ] 2.7 Update docs for rollout and compatibility rules. Objetivo: documentar dual-read, no migration requirement y la regla “never upgrade observability failure into terminal resolution”. Toca: `docs/replace-suggestion-identity-proposal.md`, `docs/review-domain-and-track-changes.md`.
