import type { PipelineContext } from "../domain/pipeline/PipelineContext";
import type { PipelineObserver } from "../domain/pipeline/PipelineEvents.types";
import { PipelineEventEmitter } from "../domain/pipeline/PipelineEvents";
import { DEFAULT_MAX_CHUNK_SIZE, POLL_INTERVAL_MS } from "../infrastructure/config";
import { hideResultsPanel } from "./ResultsPanelStore";
import { buildApplyStatusMessage, renderResultsPanel } from "./SuggestionCardRenderer";
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
import {
  collectAnalysisSuggestions,
  describeAnalysisFailures,
  describeNoSuggestionsReason,
} from "../domain/analysis/Analysis.helpers";
import type { AnalysisSessionSnapshot, AnalysisSnapshot } from "../domain/analysis/Analysis.types";
import { AnalysisPoller } from "../domain/analysis/AnalysisPoller";
import type { AnalysisPollOutcome } from "../domain/analysis/AnalysisPoller.types";

/** Applies the suggestions a retry query recovered and republishes the results surface. */
async function applyRetriedSuggestions(
  runtime: TaskpaneAnalysisHandlersRuntime,
  snapshot: AnalysisSnapshot,
  isSelection: boolean
): Promise<void> {
  const recoveredSuggestions = collectAnalysisSuggestions(snapshot);
  const applyResult = await runtime.documentPort.applySuggestions(recoveredSuggestions);
  clearTaskpaneAnalysisError();
  renderResultsPanel(
    recoveredSuggestions,
    applyResult,
    describeAnalysisFailures(snapshot),
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
function getAbortRetryKind(ctx: PipelineContext): TaskpaneAnalysisRetryKind | undefined {
  if (ctx.retryableAnalysisId) {
    return "retry-query";
  }

  if ((ctx.chunkErrors?.length ?? 0) > 0 || ctx.backendUnavailable) {
    return "full-retry";
  }

  return undefined;
}

/** Publishes pipeline-owned analysis references into the taskpane session store. */
function publishAsyncSessionSnapshot(snapshot: AnalysisSessionSnapshot): void {
  setTaskpaneAsyncAnalysisSnapshot(snapshot);
  syncTaskpaneAsyncAnalysisProgressSession();
}

/** Reads the pipeline's analysis references into the frontend async-session shape. */
function readAsyncRunSessionSnapshot(ctx: PipelineContext): AnalysisSessionSnapshot {
  return {
    isSelection: ctx.isSelection ?? false,
    activeAnalysisId: ctx.activeAnalysisId ?? null,
    retryableAnalysisId: ctx.retryableAnalysisId ?? null,
  };
}

/** Cancels the analysis currently tracked by the taskpane progress session. */
export async function handleCancelAnalysis(
  runtime: TaskpaneAnalysisHandlersRuntime
): Promise<void> {
  const session = getTaskpaneAsyncAnalysisSessionState();
  if (session.activeAnalysisId === null) {
    return;
  }

  clearTaskpaneAnalysisError();
  setTaskpaneAsyncAnalysisCanceling();
  runtime.cancelState.value = true;
  syncTaskpaneAsyncAnalysisProgressSession();

  let cancelError: string | undefined;
  try {
    await runtime.analysisPort.cancelAnalysis(session.activeAnalysisId);
  } catch (error) {
    cancelError = toTaskpaneUserMessage(error);
  }

  clearTaskpaneAsyncAnalysisSession();
  syncTaskpaneAsyncAnalysisProgressSession();
  hideTaskpaneProgress();
  setTaskpaneAnalyzeLoading(false);
  runtime.stateMachine.reset();

  if (cancelError) {
    showTaskpaneStatus(`No se pudo cancelar el análisis: ${cancelError}`, "error");
    return;
  }

  showTaskpaneStatus("Análisis cancelado en backend.", "success");
  hideResultsPanel();
}

/** Reads a retryable analysis again and waits for it to finish, without starting a new one. */
async function queryRetryableAnalysis(
  runtime: TaskpaneAnalysisHandlersRuntime,
  analysisId: string
): Promise<AnalysisPollOutcome> {
  let snapshot: AnalysisSnapshot;
  try {
    snapshot = await runtime.analysisPort.getAnalysis(analysisId);
  } catch (error) {
    return { kind: "poll-failed", message: toTaskpaneUserMessage(error) };
  }

  return new AnalysisPoller(runtime.analysisPort, POLL_INTERVAL_MS).waitUntilFinished(snapshot, {
    shouldCancel: () => runtime.cancelState.value,
  });
}

/** Publishes a retry query that failed locally again, keeping the retry affordance. */
function showRetryQueryFailedAgain(
  isSelection: boolean,
  analysisId: string,
  message: string
): void {
  publishAsyncSessionSnapshot({
    isSelection,
    activeAnalysisId: null,
    retryableAnalysisId: analysisId,
  });
  clearTaskpaneProgress();
  const reason = `La consulta volvió a fallar localmente: ${message}. Puedes reintentar otra vez.`;
  showTaskpaneAnalysisError(reason, "retry-query");
  showTaskpaneStatus(reason, "error");
}

/** Re-runs only backend polling for the analysis whose polling failed locally. */
export async function handleRetryAnalysisQuery(
  runtime: TaskpaneAnalysisHandlersRuntime
): Promise<void> {
  const session = getTaskpaneAsyncAnalysisSessionState();
  const analysisId = session.retryableAnalysisId;
  if (analysisId === null) {
    return;
  }

  clearTaskpaneAnalysisError();
  setTaskpaneAnalyzeLoading(true);
  runtime.cancelState.value = false;
  publishAsyncSessionSnapshot({
    isSelection: session.isSelection,
    activeAnalysisId: analysisId,
    retryableAnalysisId: null,
  });
  updateTaskpaneProgress(0, 1, "Reintentando consulta del backend...");

  const outcome = await queryRetryableAnalysis(runtime, analysisId);
  setTaskpaneAnalyzeLoading(false);

  if (outcome.kind === "poll-failed") {
    showRetryQueryFailedAgain(session.isSelection, analysisId, outcome.message);
    return;
  }

  clearTaskpaneAsyncAnalysisSession();
  syncTaskpaneAsyncAnalysisProgressSession();
  hideTaskpaneProgress();

  if (outcome.kind === "canceled-locally" || outcome.snapshot.status === "canceled") {
    showTaskpaneStatus("Análisis cancelado por el usuario.", "success");
    return;
  }

  if (collectAnalysisSuggestions(outcome.snapshot).length > 0) {
    await applyRetriedSuggestions(runtime, outcome.snapshot, session.isSelection);
    return;
  }

  const failures = describeAnalysisFailures(outcome.snapshot);
  const reason = describeNoSuggestionsReason(failures);
  if (failures.length > 0) {
    clearTaskpaneProgress();
    showTaskpaneAnalysisError(reason, "full-retry");
    showTaskpaneStatus(reason, "error");
    return;
  }

  showTaskpaneStatus(reason, "success");
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

      const retryKind = getAbortRetryKind(ctx);
      if (retryKind) {
        showTaskpaneAnalysisError(reason, retryKind);
      } else {
        clearTaskpaneAnalysisError();
      }

      const hasAbortError = Boolean(retryKind) || (ctx.chunkErrors?.length ?? 0) > 0;
      const statusType: "success" | "error" = hasAbortError ? "error" : "success";

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
