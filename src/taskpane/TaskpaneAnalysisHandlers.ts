import type { PipelineContext } from "../domain/pipeline/PipelineContext";
import type { PipelineObserver } from "../domain/pipeline/PipelineEvents.types";
import { PipelineEventEmitter } from "../domain/pipeline/PipelineEvents";
import { DEFAULT_MAX_CHUNK_SIZE } from "../infrastructure/config";
import { hideResultsPanel } from "./ResultsPanelStore";
import { buildApplyStatusMessage, renderResultsPanel } from "./SuggestionCardRenderer";
import { BACKEND_UNAVAILABLE_ABORT_REASON } from "./TaskpaneAnalysis.constants";
import type { TaskpaneAnalysisHandlersRuntime } from "./TaskpaneAnalysisHandlers.types";
import {
  clearTaskpaneAsyncAnalysisSession,
  getTaskpaneAsyncAnalysisSessionState,
  setTaskpaneAsyncAnalysisCanceling,
  setTaskpaneAsyncAnalysisSnapshot,
  startTaskpaneAsyncAnalysisSession,
} from "./TaskpaneAsyncAnalysisSession";
import {
  clearTaskpaneAnalysisError,
  clearTaskpaneProgress,
  hideTaskpaneProgress,
  setTaskpaneAnalyzeLoading,
  showTaskpaneAnalysisError,
  showTaskpaneStatus,
  syncTaskpaneAsyncAnalysisProgressSession,
  updateTaskpaneProgress,
} from "./TaskpaneShellStore";
import type { TaskpaneAnalysisRetryKind } from "./TaskpaneShellStore.types";
import { toTaskpaneUserMessage } from "./TaskpaneError.helpers";
import type {
  AnalysisRunSessionSnapshot,
  ChunkPollResult,
  ChunkRunReference,
} from "../domain/mastra/MastraWorkflow.types";

/** Applies successful retry-query results and republishes the results surface. */
async function applyRetriedSuggestions(
  runtime: TaskpaneAnalysisHandlersRuntime,
  successfulPolls: ChunkPollResult[],
  terminalErrors: ChunkPollResult[],
  isSelection: boolean
): Promise<void> {
  const recoveredSuggestions = successfulPolls.flatMap((result) => result.suggestions);
  if (recoveredSuggestions.length === 0) {
    showTaskpaneStatus("La consulta se recuperó, pero no devolvió sugerencias nuevas.", "success");
    return;
  }

  const applyResult = await runtime.documentPort.applySuggestions(recoveredSuggestions);
  clearTaskpaneAnalysisError();
  renderResultsPanel(
    recoveredSuggestions,
    applyResult,
    terminalErrors.flatMap((result) => (result.error ? [result.error] : [])),
    isSelection,
    runtime.cardRendererDeps
  );
  showTaskpaneStatus(
    buildApplyStatusMessage(applyResult, isSelection),
    applyResult.successCount > 0 ? "success" : "error"
  );
  void runtime.refreshCleanupVisibility();
}

/** Maps pipeline abort reasons to the retry affordance the taskpane should expose. */
function getAbortRetryKind(
  ctx: PipelineContext,
  reason: string
): TaskpaneAnalysisRetryKind | undefined {
  if ((ctx.retryableRunReferences?.length ?? 0) > 0) {
    return "retry-query";
  }

  if ((ctx.chunkErrors?.length ?? 0) > 0 || reason === BACKEND_UNAVAILABLE_ABORT_REASON) {
    return "full-retry";
  }

  return undefined;
}

/** Publishes pipeline-owned backend run references into the taskpane session store. */
function publishAsyncSessionSnapshot(snapshot: AnalysisRunSessionSnapshot): void {
  setTaskpaneAsyncAnalysisSnapshot(snapshot);
  syncTaskpaneAsyncAnalysisProgressSession();
}

/** Reads the active pipeline run references into the frontend async-session shape. */
function readAsyncRunSessionSnapshot(ctx: PipelineContext): AnalysisRunSessionSnapshot {
  return {
    isSelection: ctx.isSelection ?? false,
    activeRuns: ctx.activeRunReferences ?? [],
    retryableRuns: ctx.retryableRunReferences ?? [],
  };
}

/** Cancels every backend run currently tracked by the taskpane progress session. */
export async function handleCancelAnalysis(
  runtime: TaskpaneAnalysisHandlersRuntime
): Promise<void> {
  const session = getTaskpaneAsyncAnalysisSessionState();
  if (session.activeRuns.length === 0) {
    return;
  }

  clearTaskpaneAnalysisError();
  setTaskpaneAsyncAnalysisCanceling();
  runtime.cancelState.value = true;
  syncTaskpaneAsyncAnalysisProgressSession();

  const results = await Promise.all(
    session.activeRuns.map((run) =>
      runtime.analysisPort.cancelChunkAnalysis(run.chunkIndex, run.runId)
    )
  );
  const failedCancellations = results.filter((result) => !result.canceled);

  clearTaskpaneAsyncAnalysisSession();
  syncTaskpaneAsyncAnalysisProgressSession();
  hideTaskpaneProgress();
  setTaskpaneAnalyzeLoading(false);
  runtime.stateMachine.reset();

  if (failedCancellations.length > 0) {
    showTaskpaneStatus(
      `No se pudo cancelar ${failedCancellations.length} run(s) activo(s).`,
      "error"
    );
    return;
  }

  showTaskpaneStatus("Análisis cancelado en backend.", "success");
  hideResultsPanel();
}

/** Re-runs only backend polling for previously known run ids. */
export async function handleRetryAnalysisQuery(
  runtime: TaskpaneAnalysisHandlersRuntime
): Promise<void> {
  const session = getTaskpaneAsyncAnalysisSessionState();
  if (session.retryableRuns.length === 0) {
    return;
  }

  clearTaskpaneAnalysisError();
  setTaskpaneAnalyzeLoading(true);
  runtime.cancelState.value = false;
  publishAsyncSessionSnapshot({
    isSelection: session.isSelection,
    activeRuns: session.retryableRuns,
    retryableRuns: [],
  });
  updateTaskpaneProgress(0, session.retryableRuns.length, "Reintentando consulta del backend...");

  const settled = await Promise.all(
    session.retryableRuns.map((reference) => runtime.analysisPort.retryPollChunkAnalysis(reference))
  );

  const successfulPolls = settled.filter((result) => result.status === "success");
  const stillRetryable = settled
    .filter((result) => result.status === "retryable-failure")
    .map<ChunkRunReference>((result) => ({ chunkIndex: result.chunkIndex, runId: result.runId }));
  const terminalErrors = settled.filter(
    (result) => result.status !== "success" && result.status !== "retryable-failure"
  );

  if (successfulPolls.length > 0) {
    await applyRetriedSuggestions(runtime, successfulPolls, terminalErrors, session.isSelection);
  }

  publishAsyncSessionSnapshot({
    isSelection: session.isSelection,
    activeRuns: [],
    retryableRuns: stillRetryable,
  });
  setTaskpaneAnalyzeLoading(false);
  syncTaskpaneAsyncAnalysisProgressSession();

  if (stillRetryable.length > 0) {
    clearTaskpaneProgress();
    showTaskpaneAnalysisError(
      "La consulta volvió a fallar localmente. Puedes reintentar otra vez.",
      "retry-query"
    );
    showTaskpaneStatus(
      "La consulta volvió a fallar localmente. Puedes reintentar otra vez.",
      "error"
    );
    return;
  }

  if (terminalErrors.length > 0) {
    if (successfulPolls.length === 0) {
      clearTaskpaneProgress();
      showTaskpaneAnalysisError(
        `${terminalErrors.length} run(s) terminaron con estado terminal en backend.`,
        "full-retry"
      );
    }
    showTaskpaneStatus(
      `${terminalErrors.length} run(s) terminaron con estado terminal en backend.`,
      "error"
    );
    return;
  }

  if (successfulPolls.length > 0) {
    clearTaskpaneAnalysisError();
    clearTaskpaneAsyncAnalysisSession();
    syncTaskpaneAsyncAnalysisProgressSession();
    hideTaskpaneProgress();
  }
}

/** Runs the main analysis pipeline and republishes state through the taskpane stores. */
export async function handleAnalyze(runtime: TaskpaneAnalysisHandlersRuntime): Promise<void> {
  if (runtime.stateMachine.isRunning) {
    console.warn("⚠️ [Taskpane] Pipeline ya en ejecución — ignorando click");
    return;
  }

  clearTaskpaneAnalysisError();
  setTaskpaneAnalyzeLoading(true);
  runtime.cancelState.value = false;
  startTaskpaneAsyncAnalysisSession();
  syncTaskpaneAsyncAnalysisProgressSession();
  hideResultsPanel();

  const emitter = new PipelineEventEmitter();
  const ctx: PipelineContext = {
    documentPort: runtime.documentPort,
    analysisPort: runtime.analysisPort,
    emitter,
    genero: runtime.getSelectedGenero(),
    maxChunkSize: DEFAULT_MAX_CHUNK_SIZE,
    shouldCancelAnalysis: () => runtime.cancelState.value,
  };

  const uiObserver: PipelineObserver = {
    onPhaseStart(_phase, message) {
      updateTaskpaneProgress(0, 1, message);
    },
    onProgress(current, total, message) {
      updateTaskpaneProgress(current, total, message);
      publishAsyncSessionSnapshot(readAsyncRunSessionSnapshot(ctx));
    },
    onAbort(reason) {
      publishAsyncSessionSnapshot(readAsyncRunSessionSnapshot(ctx));
      clearTaskpaneProgress();

      const retryKind = getAbortRetryKind(ctx, reason);
      if (retryKind) {
        showTaskpaneAnalysisError(reason, retryKind);
      } else {
        clearTaskpaneAnalysisError();
      }

      const statusType = retryKind
        ? "error"
        : ctx.shouldCancelAnalysis?.()
          ? "success"
          : (ctx.chunkErrors?.length ?? 0) > 0
            ? "error"
            : "success";
      showTaskpaneStatus(reason, statusType);
    },
    onComplete(suggestions, result, chunkErrors, isSelection) {
      publishAsyncSessionSnapshot(readAsyncRunSessionSnapshot(ctx));
      hideTaskpaneProgress();
      clearTaskpaneAnalysisError();
      renderResultsPanel(suggestions, result, chunkErrors, isSelection, runtime.cardRendererDeps);
      void runtime.refreshCleanupVisibility();
      showTaskpaneStatus(
        buildApplyStatusMessage(result, isSelection),
        result.successCount > 0 ? "success" : "error"
      );
    },
  };

  emitter.subscribe(uiObserver);
  runtime.stateMachine.transition("reading");

  try {
    console.log("🚀 [Taskpane] Pipeline iniciado");
    await runtime.orchestrator.run(ctx);
    console.log(`✅ [Taskpane] Pipeline completado. Abortado: ${ctx.aborted ?? false}`);
  } catch (error) {
    console.error("💥 [Taskpane] Error no capturado en pipeline:", error);
    clearTaskpaneProgress();
    const message = toTaskpaneUserMessage(error);
    showTaskpaneAnalysisError(message, "full-retry");
    showTaskpaneStatus(message, "error");
  } finally {
    publishAsyncSessionSnapshot(readAsyncRunSessionSnapshot(ctx));
    runtime.stateMachine.reset();
    runtime.cancelState.value = false;
    setTaskpaneAnalyzeLoading(false);
    emitter.clear();
  }
}
