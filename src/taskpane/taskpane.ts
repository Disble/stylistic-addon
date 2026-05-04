/**
 * Task-pane — Composition Root and top-level event handlers.
 *
 * Responsibilities:
 * - Instantiate all adapters, decorators, mediators, and the pipeline.
 * - Initialize after the React shell is mounted and Office confirms the host is Word.
 * - Bind top-level DOM event handlers (analyze, cleanup, disable TC).
 * - Register a `PipelineObserver` to relay pipeline events into presentation stores.
 *
 * React now owns shell/results rendering. This module still owns orchestration,
 * top-level host event binding, and publication into presentation facades.
 *
 * @module taskpane
 */

import { FeedbackAdapter } from "../adapters/mastra/FeedbackAdapter";
import { MastraAdapter } from "../adapters/mastra/MastraAdapter";
import { ConsoleResolutionObservabilityAdapter } from "../adapters/observability/ConsoleResolutionObservabilityAdapter";
import { RetryAnalysisDecorator } from "../adapters/RetryAnalysisDecorator";
import { WordAdapter } from "../adapters/word/WordAdapter";
import { AnalyzeChunksHandler } from "../domain/pipeline/handlers/AnalyzeChunksHandler";
import { ApplySuggestionsHandler } from "../domain/pipeline/handlers/ApplySuggestionsHandler";
import { CheckConnectionHandler } from "../domain/pipeline/handlers/CheckConnectionHandler";
import { ChunkTextHandler } from "../domain/pipeline/handlers/ChunkTextHandler";
import { DeduplicateHandler } from "../domain/pipeline/handlers/DeduplicateHandler";
import { GuardAppliedHandler } from "../domain/pipeline/handlers/GuardAppliedHandler";
import { ReadTextHandler } from "../domain/pipeline/handlers/ReadTextHandler";
import type { PipelineContext } from "../domain/pipeline/PipelineContext";
import { PipelineEventEmitter, type PipelineObserver } from "../domain/pipeline/PipelineEvents";
import { PipelineOrchestrator } from "../domain/pipeline/PipelineOrchestrator";
import { PipelineStateMachine } from "../domain/pipeline/PipelineStateMachine";
import type { IFeedbackPort } from "../domain/ports";
import { ReviewSessionMediator } from "../domain/review/ReviewSessionMediator";
import { DEFAULT_MAX_CHUNK_SIZE, MAX_RETRIES, RETRY_BASE_DELAY_MS } from "../infrastructure/config";
import {
  buildApplyStatusMessage,
  type ResultsPanelDeps,
  renderResultsPanel,
} from "./SuggestionCardRenderer";
import {
  getRequiredElement,
  getSelectedGenero,
  hideProgress,
  setCleanupCtaVisible,
  setAnalyzeLoading,
  setDisableTrackChangesCtaVisible,
  showStatus,
  toUserMessage,
  updateProgress,
} from "./TaskpaneUi";

type OfficeLike = {
  HostType?: {
    Word?: string | Office.HostType;
  };
};

type DocumentLike = Pick<Document, "getElementById">;

// ---------------------------------------------------------------------------
// Infrastructure — built once, reused across pipeline runs
// ---------------------------------------------------------------------------

const observabilityPort = new ConsoleResolutionObservabilityAdapter();
const documentPort = new WordAdapter(undefined, observabilityPort);
const analysisPort = new RetryAnalysisDecorator(
  new MastraAdapter(),
  MAX_RETRIES,
  RETRY_BASE_DELAY_MS
);

const feedbackPort: IFeedbackPort = new FeedbackAdapter();
const reviewSessionMediator = new ReviewSessionMediator(documentPort, feedbackPort);

const orchestrator = new PipelineOrchestrator([
  new ReadTextHandler(),
  new CheckConnectionHandler(),
  new ChunkTextHandler(),
  new AnalyzeChunksHandler(),
  new DeduplicateHandler(),
  new GuardAppliedHandler(),
  new ApplySuggestionsHandler(),
]);

const stateMachine = new PipelineStateMachine();

/** Deps injected into the card renderer — closures over module-level ports. */
const cardRendererDeps: ResultsPanelDeps = {
  navigateToText: (target) => documentPort.navigateToText(target),
  acceptSuggestion: (s, comment) => reviewSessionMediator.acceptSuggestion(s, comment),
  rejectSuggestion: (s, comment) => reviewSessionMediator.rejectSuggestion(s, comment),
};

// ---------------------------------------------------------------------------
// Initialization
// ---------------------------------------------------------------------------

/**
 * Initializes taskpane DOM bindings once the React shell already exists.
 * Hides the sideload message, shows the app body, and binds event handlers.
 */
export function bootstrapTaskpane(
  doc: DocumentLike | undefined = globalThis.document,
  office: OfficeLike | undefined = globalThis.Office as unknown as OfficeLike | undefined
): void {
  if (!doc?.getElementById) {
    return;
  }

  const sideloadMessage = doc.getElementById("sideload-msg");
  const appBody = doc.getElementById("app-body");
  const analyzeButton = doc.getElementById("btn-analyze") as HTMLButtonElement | null;
  const cleanupButton = doc.getElementById("btn-cleanup") as HTMLButtonElement | null;
  const disableTrackChangesButton = doc.getElementById(
    "btn-disable-track-changes"
  ) as HTMLButtonElement | null;

  if (
    !(sideloadMessage && appBody && analyzeButton && cleanupButton && disableTrackChangesButton)
  ) {
    return;
  }

  const wordHost = String(office?.HostType?.Word ?? "Word");
  sideloadMessage.dataset.officeHost = wordHost;
  sideloadMessage.style.display = "none";
  appBody.style.display = "flex";
  analyzeButton.onclick = handleAnalyze;
  cleanupButton.onclick = handleCleanup;
  disableTrackChangesButton.onclick = handleDisableTrackChanges;
  void refreshCleanupVisibility();
  void refreshTrackChangesCtaVisibility();
}

// ---------------------------------------------------------------------------
// Refresh helpers
// ---------------------------------------------------------------------------

/**
 * Syncs the cleanup CTA visibility with the current document state.
 * Shows the section only when there are deletable Stylistic comments.
 */
async function refreshCleanupVisibility(): Promise<void> {
  try {
    const { deletable } = await documentPort.getCleanupPreview();
    setCleanupCtaVisible(deletable > 0);
  } catch (error) {
    console.warn("⚠️ [Taskpane] No se pudo calcular la visibilidad de limpieza:", error);
  }
}

/**
 * Rehydrates Track Changes CTA visibility from the authoritative document state.
 */
async function refreshTrackChangesCtaVisibility(): Promise<void> {
  try {
    const taskpaneState = await reviewSessionMediator.rehydrateTaskpaneState();
    setDisableTrackChangesCtaVisible(taskpaneState.showDisableTrackChangesCta);
  } catch (error) {
    console.warn(
      "⚠️ [Taskpane] No se pudo calcular la visibilidad del CTA de Track Changes:",
      error
    );
  }
}

// ---------------------------------------------------------------------------
// Main Event Handler — Analysis Pipeline
// ---------------------------------------------------------------------------

/**
 * Handles the "Analizar y sugerir" button click.
 *
 * Creates the `PipelineContext` with injected adapters and a UI observer,
 * then delegates entirely to `PipelineOrchestrator.run()`.
 * All UI updates come back via `PipelineEventEmitter` (Observer pattern).
 */
async function handleAnalyze(): Promise<void> {
  if (stateMachine.isRunning) {
    console.warn("⚠️ [Taskpane] Pipeline ya en ejecución — ignorando click");
    return;
  }

  setAnalyzeLoading(true);
  getRequiredElement("results-panel").style.display = "none";

  const emitter = new PipelineEventEmitter();
  const ctx: PipelineContext = {
    documentPort,
    analysisPort,
    emitter,
    genero: getSelectedGenero(),
    maxChunkSize: DEFAULT_MAX_CHUNK_SIZE,
  };

  const uiObserver: PipelineObserver = {
    onPhaseStart(_phase, message) {
      updateProgress(0, 1, message);
    },
    onProgress(current, total, message) {
      updateProgress(current, total, message);
    },
    onAbort(reason) {
      hideProgress();
      showStatus(reason, (ctx.chunkErrors?.length ?? 0) > 0 ? "error" : "success");
    },
    onComplete(suggestions, result, chunkErrors, isSelection) {
      hideProgress();
      renderResultsPanel(suggestions, result, chunkErrors, isSelection, cardRendererDeps);
      void refreshCleanupVisibility();
      showStatus(
        buildApplyStatusMessage(result, isSelection),
        result.successCount > 0 ? "success" : "error"
      );
    },
  };

  emitter.subscribe(uiObserver);
  stateMachine.transition("reading");

  try {
    console.log("🚀 [Taskpane] Pipeline iniciado");
    await orchestrator.run(ctx);
    console.log(`✅ [Taskpane] Pipeline completado. Abortado: ${ctx.aborted ?? false}`);
  } catch (error) {
    console.error("💥 [Taskpane] Error no capturado en pipeline:", error);
    hideProgress();
    showStatus(toUserMessage(error), "error");
  } finally {
    stateMachine.reset();
    setAnalyzeLoading(false);
    emitter.clear();
  }
}

// ---------------------------------------------------------------------------
// Comment Cleanup Handler
// ---------------------------------------------------------------------------

async function handleCleanup(): Promise<void> {
  console.log("🧽 [Taskpane] Iniciando limpieza de comentarios resueltos...");
  const btn = document.getElementById("btn-cleanup") as HTMLButtonElement;
  const label = getRequiredElement("btn-cleanup-label");

  btn.disabled = true;
  label.textContent = "Limpiando...";

  try {
    const { deleted, kept } = await documentPort.cleanupResolvedComments();
    console.log(`🧽 [Taskpane] Limpieza: ${deleted} eliminados, ${kept} conservados`);
    showStatus(`${deleted} comentario(s) eliminado(s), ${kept} conservado(s).`, "success");
    setCleanupCtaVisible(kept > 0);
  } catch (error) {
    showStatus(toUserMessage(error), "error");
  } finally {
    btn.disabled = false;
    label.textContent = "Limpiar comentarios resueltos";
  }
}

// ---------------------------------------------------------------------------
// Disable Track Changes Handler
// ---------------------------------------------------------------------------

async function handleDisableTrackChanges(): Promise<void> {
  const btn = document.getElementById("btn-disable-track-changes") as HTMLButtonElement | null;
  const label = document.getElementById("btn-disable-track-changes-label");

  if (!(btn && label)) {
    return;
  }

  btn.disabled = true;
  label.textContent = "Desactivando...";

  try {
    const taskpaneState = await reviewSessionMediator.disableTrackChanges();
    setDisableTrackChangesCtaVisible(taskpaneState.showDisableTrackChangesCta);
    showStatus("Control de cambios desactivado.", "success");
  } catch (error) {
    showStatus(toUserMessage(error), "error");
  } finally {
    btn.disabled = false;
    label.textContent = "Desactivar control de cambios";
  }
}
