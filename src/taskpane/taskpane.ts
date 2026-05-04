/**
 * Task-pane — Composition Root and top-level event handlers.
 *
 * Responsibilities:
 * - Instantiate all adapters, decorators, mediators, and the pipeline.
 * - Initialize after the React shell is mounted and Office confirms the host is Word.
 * - Expose top-level workflow handlers to the React shell.
 * - Register a `PipelineObserver` to relay pipeline events into presentation stores.
 *
 * React now owns shell/results rendering and UI event wiring. This module still owns
 * orchestration and publication into presentation stores/facades.
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
import { hideResultsPanel } from "./ResultsPanelStore";
import {
  buildApplyStatusMessage,
  type ResultsPanelDeps,
  renderResultsPanel,
} from "./SuggestionCardRenderer";
import {
  getTaskpaneShellState,
  hideTaskpaneProgress,
  setTaskpaneAnalyzeLoading,
  setTaskpaneCleanupVisible,
  setTaskpaneCleanupLoading,
  setTaskpaneDisableTrackChangesCtaVisible,
  setTaskpaneDisableTrackChangesLoading,
  showTaskpaneStatus,
  updateTaskpaneProgress,
} from "./TaskpaneShellStore";

function toUserMessage(error: unknown): string {
  if (!(error instanceof Error)) {
    if (typeof error === "string") {
      return error;
    }
    try {
      return JSON.stringify(error);
    } catch {
      return "Error no serializable";
    }
  }

  const officeError = error as Error & { code?: string };
  switch (officeError.code) {
    case "AccessDenied":
      return "El documento está protegido o es de solo lectura.";
    case "InvalidArgument":
      return "Argumento inválido al comunicarse con Word.";
    case "ItemNotFound":
      return "No se encontró el elemento solicitado en el documento.";
    default:
      return officeError.message;
  }
}

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
 * Rehydrates shell visibility from document state after the React shell mounts.
 */
export function bootstrapTaskpane(): void {
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
    setTaskpaneCleanupVisible(deletable > 0);
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
    setTaskpaneDisableTrackChangesCtaVisible(taskpaneState.showDisableTrackChangesCta);
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
export async function handleAnalyze(): Promise<void> {
  if (stateMachine.isRunning) {
    console.warn("⚠️ [Taskpane] Pipeline ya en ejecución — ignorando click");
    return;
  }

  setTaskpaneAnalyzeLoading(true);
  hideResultsPanel();

  const emitter = new PipelineEventEmitter();
  const ctx: PipelineContext = {
    documentPort,
    analysisPort,
    emitter,
    genero: getTaskpaneShellState().selectedGenero,
    maxChunkSize: DEFAULT_MAX_CHUNK_SIZE,
  };

  const uiObserver: PipelineObserver = {
    onPhaseStart(_phase, message) {
      updateTaskpaneProgress(0, 1, message);
    },
    onProgress(current, total, message) {
      updateTaskpaneProgress(current, total, message);
    },
    onAbort(reason) {
      hideTaskpaneProgress();
      showTaskpaneStatus(reason, (ctx.chunkErrors?.length ?? 0) > 0 ? "error" : "success");
    },
    onComplete(suggestions, result, chunkErrors, isSelection) {
      hideTaskpaneProgress();
      renderResultsPanel(suggestions, result, chunkErrors, isSelection, cardRendererDeps);
      void refreshCleanupVisibility();
      showTaskpaneStatus(
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
    hideTaskpaneProgress();
    showTaskpaneStatus(toUserMessage(error), "error");
  } finally {
    stateMachine.reset();
    setTaskpaneAnalyzeLoading(false);
    emitter.clear();
  }
}

// ---------------------------------------------------------------------------
// Comment Cleanup Handler
// ---------------------------------------------------------------------------

export async function handleCleanup(): Promise<void> {
  console.log("🧽 [Taskpane] Iniciando limpieza de comentarios resueltos...");
  if (getTaskpaneShellState().isCleanupLoading) {
    return;
  }

  setTaskpaneCleanupLoading(true);

  try {
    const { deleted, kept } = await documentPort.cleanupResolvedComments();
    console.log(`🧽 [Taskpane] Limpieza: ${deleted} eliminados, ${kept} conservados`);
    showTaskpaneStatus(`${deleted} comentario(s) eliminado(s), ${kept} conservado(s).`, "success");
    setTaskpaneCleanupVisible(kept > 0);
  } catch (error) {
    showTaskpaneStatus(toUserMessage(error), "error");
  } finally {
    setTaskpaneCleanupLoading(false);
  }
}

// ---------------------------------------------------------------------------
// Disable Track Changes Handler
// ---------------------------------------------------------------------------

export async function handleDisableTrackChanges(): Promise<void> {
  if (getTaskpaneShellState().isDisableTrackChangesLoading) {
    return;
  }

  setTaskpaneDisableTrackChangesLoading(true);

  try {
    const taskpaneState = await reviewSessionMediator.disableTrackChanges();
    setTaskpaneDisableTrackChangesCtaVisible(taskpaneState.showDisableTrackChangesCta);
    showTaskpaneStatus("Control de cambios desactivado.", "success");
  } catch (error) {
    showTaskpaneStatus(toUserMessage(error), "error");
  } finally {
    setTaskpaneDisableTrackChangesLoading(false);
  }
}
