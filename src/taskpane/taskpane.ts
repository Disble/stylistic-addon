/* global document, Office, console, setTimeout, HTMLButtonElement, HTMLSelectElement */

/**
 * Task-pane UI — event binding, progress rendering, and results display.
 *
 * Responsibilities (exclusively UI):
 * - Initialize when Office.js confirms the host is Word.
 * - Bind DOM event handlers to UI controls.
 * - Register a `PipelineObserver` to update the UI during analysis.
 * - Render results after pipeline completion.
 * - Handle the comment cleanup button (direct adapter call, no pipeline needed).
 *
 * This module contains NO pipeline logic, NO Word API calls, and NO backend
 * communication. All analysis flows through the `PipelineOrchestrator`.
 * All document operations go through `WordAdapter` via `IDocumentPort`.
 *
 * @module taskpane
 */

import { PipelineOrchestrator } from "../domain/pipeline/PipelineOrchestrator";
import { PipelineStateMachine } from "../domain/pipeline/PipelineStateMachine";
import { PipelineEventEmitter, PipelineObserver } from "../domain/pipeline/PipelineEvents";
import { PipelineContext } from "../domain/pipeline/PipelineContext";

import { ReadTextHandler } from "../domain/pipeline/handlers/ReadTextHandler";
import { CheckConnectionHandler } from "../domain/pipeline/handlers/CheckConnectionHandler";
import { ChunkTextHandler } from "../domain/pipeline/handlers/ChunkTextHandler";
import { AnalyzeChunksHandler } from "../domain/pipeline/handlers/AnalyzeChunksHandler";
import { DeduplicateHandler } from "../domain/pipeline/handlers/DeduplicateHandler";
import { GuardAppliedHandler } from "../domain/pipeline/handlers/GuardAppliedHandler";
import { ApplySuggestionsHandler } from "../domain/pipeline/handlers/ApplySuggestionsHandler";

import { WordAdapter } from "../adapters/word/WordAdapter";
import { MastraAdapter } from "../adapters/mastra/MastraAdapter";
import { RetryAnalysisDecorator } from "../adapters/RetryAnalysisDecorator";

import { Suggestion, InsertionResult } from "../domain/types";
import { DEFAULT_MAX_CHUNK_SIZE, MAX_RETRIES, RETRY_BASE_DELAY_MS } from "../infrastructure/config";

/** Duration (ms) before the status bar message auto-hides. */
const STATUS_DISPLAY_MS = 4000;

// ---------------------------------------------------------------------------
// Infrastructure — built once, reused across pipeline runs
// ---------------------------------------------------------------------------

const documentPort = new WordAdapter();
const analysisPort = new RetryAnalysisDecorator(
  new MastraAdapter(),
  MAX_RETRIES,
  RETRY_BASE_DELAY_MS
);

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

// ---------------------------------------------------------------------------
// Initialization
// ---------------------------------------------------------------------------

/**
 * Entry point — called once Office.js confirms the host is Word.
 * Hides the sideload message, shows the app body, and binds event handlers.
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";
    document.getElementById("btn-analyze")!.onclick = handleAnalyze;
    document.getElementById("btn-cleanup")!.onclick = handleCleanup;
  }
});

// ---------------------------------------------------------------------------
// UI Helpers
// ---------------------------------------------------------------------------

/**
 * Displays a temporary status message at the bottom of the task pane.
 * Auto-hides after {@link STATUS_DISPLAY_MS} milliseconds.
 */
function showStatus(message: string, type: "success" | "error"): void {
  const bar = document.getElementById("status-bar")!;
  bar.textContent = message;
  bar.className = `stylistic-status ${type}`;
  bar.style.display = "block";
  setTimeout(() => {
    bar.style.display = "none";
  }, STATUS_DISPLAY_MS);
}

/**
 * Toggles the "Analizar y sugerir" button between normal and loading states.
 */
function setAnalyzeLoading(loading: boolean): void {
  const btn = document.getElementById("btn-analyze") as HTMLButtonElement;
  const label = document.getElementById("btn-analyze-label")!;
  const select = document.getElementById("profile-select") as HTMLSelectElement;
  btn.disabled = loading;
  select.disabled = loading;
  label.textContent = loading ? "Analizando..." : "Analizar y sugerir";
}

/**
 * Updates the progress bar and text in the progress area.
 */
function updateProgress(current: number, total: number, message: string): void {
  const container = document.getElementById("progress-container")!;
  const bar = document.getElementById("progress-bar")!;
  const text = document.getElementById("progress-text")!;

  container.style.display = "block";
  const percentage = total > 0 ? Math.round((current / total) * 100) : 0;
  bar.style.width = `${percentage}%`;
  text.textContent = message;
}

/**
 * Hides the progress bar (called on pipeline completion or abort).
 */
function hideProgress(): void {
  setTimeout(() => {
    const container = document.getElementById("progress-container");
    if (container) container.style.display = "none";
  }, 1000);
}

/**
 * Renders the results panel showing each suggestion and its outcome.
 */
function renderResults(
  suggestions: Suggestion[],
  result: InsertionResult,
  chunkErrors: string[],
  isSelection: boolean
): void {
  const panel = document.getElementById("results-panel")!;
  const summary = document.getElementById("results-summary")!;
  const list = document.getElementById("results-list")!;

  const total = suggestions.length;
  const applied = result.successCount;
  const failed = result.failedSuggestions.length;

  const scopePrefix = isSelection ? "Sobre selección — " : "";
  let summaryText = `${scopePrefix}${applied} de ${total} sugerencias aplicadas como Track Changes.`;
  if (failed > 0) summaryText += ` ${failed} no encontrada(s) en el texto.`;
  if (chunkErrors.length > 0) summaryText += ` ${chunkErrors.length} fragmento(s) con error.`;
  summary.textContent = summaryText;

  list.innerHTML = "";
  for (const s of suggestions) {
    const li = document.createElement("li");
    const isFailed = result.failedSuggestions.some((f) => f.id === s.id);

    if (isFailed) {
      li.innerHTML =
        `<span class="result-category">${escapeHtml(s.category)}</span>` +
        `<span class="result-failed">No encontrado: "${escapeHtml(s.originalText)}"</span>` +
        `<span class="result-justification">${escapeHtml(s.justification)}</span>`;
    } else {
      li.innerHTML =
        `<span class="result-category">${escapeHtml(s.category)}</span>` +
        `<span class="result-change">` +
        `<span class="result-original">${escapeHtml(s.originalText)}</span>` +
        `<span class="result-arrow">&rarr;</span>` +
        `<span class="result-suggested">${escapeHtml(s.suggestedText)}</span>` +
        `</span>` +
        `<span class="result-justification">${escapeHtml(s.justification)}</span>`;
    }
    list.appendChild(li);
  }

  panel.style.display = "block";
}

/** Escapes a string for safe insertion into innerHTML (XSS prevention). */
function escapeHtml(str: string): string {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

/**
 * Translates a caught error into a user-friendly Spanish message.
 * Recognizes common Office.js error codes.
 */
function toUserMessage(error: unknown): string {
  if (!(error instanceof Error)) return String(error);
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

/** Returns the currently selected profile ID from the dropdown. */
function getSelectedProfile(): string {
  const select = document.getElementById("profile-select") as HTMLSelectElement;
  return select.value;
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
  // Guard: prevent concurrent runs using the State Machine
  if (stateMachine.isRunning) {
    console.warn("⚠️ [Taskpane] Pipeline ya en ejecución — ignorando click");
    return;
  }

  setAnalyzeLoading(true);
  document.getElementById("results-panel")!.style.display = "none";
  document.getElementById("cleanup-section")!.style.display = "none";

  const emitter = new PipelineEventEmitter();

  // UI Observer — maps pipeline events to DOM updates
  const uiObserver: PipelineObserver = {
    onPhaseStart(_phase, message) {
      updateProgress(0, 1, message);
    },
    onProgress(current, total, message) {
      updateProgress(current, total, message);
    },
    onAbort(reason) {
      hideProgress();
      showStatus(reason, chunkErrorOccurred ? "error" : "success");
    },
    onComplete(suggestions, result, chunkErrors, isSelection) {
      hideProgress();
      renderResults(suggestions, result, chunkErrors, isSelection);

      if (result.successCount > 0) {
        document.getElementById("cleanup-section")!.style.display = "block";
      }

      const scopeSuffix = isSelection ? " (selección)" : "";
      if (result.failedSuggestions.length > 0 && result.successCount > 0) {
        showStatus(
          `${result.successCount} aplicada(s), ${result.failedSuggestions.length} no encontrada(s)${scopeSuffix}.`,
          "success"
        );
      } else if (result.successCount > 0) {
        showStatus(
          `${result.successCount} sugerencia(s) insertada(s) como Track Changes${scopeSuffix}.`,
          "success"
        );
      } else {
        showStatus("Ninguna sugerencia pudo aplicarse al documento actual.", "error");
      }
    },
  };

  let chunkErrorOccurred = false;
  emitter.subscribe(uiObserver);

  // Also track chunk errors for the abort status type
  emitter.subscribe({
    onAbort() {
      // Already handled in uiObserver; this hook sets chunkErrorOccurred
    },
  });

  const ctx: PipelineContext = {
    documentPort,
    analysisPort,
    emitter,
    profile: getSelectedProfile(),
    maxChunkSize: DEFAULT_MAX_CHUNK_SIZE,
  };

  stateMachine.transition("reading");

  try {
    console.log("🚀 [Taskpane] Pipeline iniciado");
    await orchestrator.run(ctx);

    if (ctx.chunkErrors && ctx.chunkErrors.length > 0) {
      chunkErrorOccurred = true;
    }

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

/**
 * Handles the "Limpiar comentarios resueltos" button click.
 * Direct call to `IDocumentPort.cleanupResolvedComments()` — no pipeline needed.
 */
async function handleCleanup(): Promise<void> {
  console.log("🧽 [Taskpane] Iniciando limpieza de comentarios resueltos...");
  const btn = document.getElementById("btn-cleanup") as HTMLButtonElement;
  const label = document.getElementById("btn-cleanup-label")!;

  btn.disabled = true;
  label.textContent = "Limpiando...";

  try {
    const { deleted, kept } = await documentPort.cleanupResolvedComments();
    console.log(`🧽 [Taskpane] Limpieza: ${deleted} eliminados, ${kept} conservados`);
    showStatus(`${deleted} comentario(s) eliminado(s), ${kept} conservado(s).`, "success");

    if (kept === 0) {
      document.getElementById("cleanup-section")!.style.display = "none";
    }
  } catch (error) {
    showStatus(toUserMessage(error), "error");
  } finally {
    btn.disabled = false;
    label.textContent = "Limpiar comentarios resueltos";
  }
}
