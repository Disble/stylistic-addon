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
import { FeedbackAdapter } from "../adapters/mastra/FeedbackAdapter";
import { RetryAnalysisDecorator } from "../adapters/RetryAnalysisDecorator";

import { Suggestion, InsertionResult, FeedbackPayload, SuggestionState } from "../domain/types";
import { IFeedbackPort } from "../domain/ports";
import { SuggestionStateMachine, mapResultStatusToState } from "../domain/suggestion/SuggestionStateMachine";
import { DEFAULT_MAX_CHUNK_SIZE, MAX_RETRIES, RETRY_BASE_DELAY_MS } from "../infrastructure/config";

/** Duration (ms) before the status bar message auto-hides. */
const STATUS_DISPLAY_MS = 4000;

type OfficeLike = {
  onReady(callback: (info: { host: string }) => void): Promise<unknown> | void;
  HostType?: {
    Word?: string;
  };
};

type DocumentLike = Pick<Document, "getElementById">;

// ---------------------------------------------------------------------------
// Infrastructure — built once, reused across pipeline runs
// ---------------------------------------------------------------------------

const documentPort = new WordAdapter();
const analysisPort = new RetryAnalysisDecorator(
  new MastraAdapter(),
  MAX_RETRIES,
  RETRY_BASE_DELAY_MS
);

const feedbackPort: IFeedbackPort = new FeedbackAdapter();

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
export function bootstrapTaskpane(
  office: OfficeLike | undefined = globalThis.Office as unknown as OfficeLike | undefined,
  doc: DocumentLike | undefined = globalThis.document
): void {
  if (!office?.onReady || !doc?.getElementById) {
    return;
  }

  office.onReady((info) => {
    const wordHost = office.HostType?.Word ?? "Word";
    if (info.host !== wordHost) {
      return;
    }

    const sideloadMessage = doc.getElementById("sideload-msg");
    const appBody = doc.getElementById("app-body");
    const analyzeButton = doc.getElementById("btn-analyze") as HTMLButtonElement | null;
    const cleanupButton = doc.getElementById("btn-cleanup") as HTMLButtonElement | null;

    if (!(sideloadMessage && appBody && analyzeButton && cleanupButton)) {
      return;
    }

    sideloadMessage.style.display = "none";
    appBody.style.display = "flex";
    analyzeButton.onclick = handleAnalyze;
    cleanupButton.onclick = handleCleanup;
  });
}

bootstrapTaskpane();

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
    const isCommentOnly = s.type === "comment-only";

    if (isFailed) {
      li.innerHTML =
        `<span class="result-category">${escapeHtml(s.category)}</span>` +
        `<span class="result-failed">No encontrado: "${escapeHtml(s.originalText)}"</span>` +
        `<span class="result-justification">${escapeHtml(s.justification)}</span>`;
    } else {
      // category + severity badge always shown
      li.innerHTML =
        `<span class="result-category">${escapeHtml(s.category)}</span>` +
        `<span class="result-severity result-severity--${escapeHtml(s.severity)}">${escapeHtml(s.severity)}</span>`;

      // comment-only type badge
      if (isCommentOnly) {
        const typeBadge = document.createElement("span");
        typeBadge.className = "result-type-badge result-type-badge--comment";
        typeBadge.textContent = "comentario";
        li.appendChild(typeBadge);
      }

      // diff block only for track-change suggestions
      if (!isCommentOnly) {
        const changeSpan = document.createElement("span");
        changeSpan.innerHTML =
          `<span class="result-original">${escapeHtml(s.originalText)}</span>` +
          `<span class="result-arrow">&rarr;</span>` +
          `<span class="result-suggested">${escapeHtml(s.suggestedText ?? "")}</span>`;
        changeSpan.className = "result-change";
        li.appendChild(changeSpan);
      }

      // justification always shown
      const justSpan = document.createElement("span");
      justSpan.className = "result-justification";
      justSpan.textContent = s.justification;
      li.appendChild(justSpan);

      // Build accept/reject buttons programmatically so they are testable DOM nodes
      const actionsSpan = document.createElement("span");
      actionsSpan.className = "result-actions";

      const acceptBtn = document.createElement("button");
      acceptBtn.className = isCommentOnly ? "result-action-btn result-action-btn--text" : "result-action-btn";
      acceptBtn.setAttribute("data-action", "accept");
      acceptBtn.setAttribute("data-suggestion-id", s.id);
      acceptBtn.setAttribute("aria-label", "Aceptar sugerencia");
      acceptBtn.textContent = isCommentOnly ? "Entendido" : "✓";

      const rejectBtn = document.createElement("button");
      rejectBtn.className = isCommentOnly ? "result-action-btn result-action-btn--text" : "result-action-btn";
      rejectBtn.setAttribute("data-action", "reject");
      rejectBtn.setAttribute("data-suggestion-id", s.id);
      rejectBtn.setAttribute("aria-label", "Rechazar sugerencia");
      rejectBtn.textContent = isCommentOnly ? "Ignorar" : "✗";

      const feedbackBtn = document.createElement("button");
      feedbackBtn.className = "feedback-btn";
      feedbackBtn.setAttribute("data-action", "feedback");
      feedbackBtn.setAttribute("aria-label", "Dejar feedback");
      feedbackBtn.textContent = "💬";

      actionsSpan.appendChild(acceptBtn);
      actionsSpan.appendChild(rejectBtn);
      actionsSpan.appendChild(feedbackBtn);
      li.appendChild(actionsSpan);

      // Accordion + textarea for optional comment
      const accordion = document.createElement("div");
      accordion.className = "feedback-accordion";

      const textarea = document.createElement("textarea");
      textarea.className = "feedback-textarea";
      textarea.setAttribute("placeholder", "Comentario opcional...");
      accordion.appendChild(textarea);
      li.appendChild(accordion);
    }

    list.appendChild(li);

    if (!isFailed) {
      const acceptBtn = li.querySelector("[data-action=\"accept\"]") as HTMLButtonElement | null;
      const rejectBtn = li.querySelector("[data-action=\"reject\"]") as HTMLButtonElement | null;
      const feedbackBtnEl = li.querySelector("[data-action=\"feedback\"]") as HTMLButtonElement | null;
      const accordionEl = li.querySelector(".feedback-accordion") as HTMLElement | null;

      if (feedbackBtnEl && accordionEl) {
        feedbackBtnEl.addEventListener("click", () => {
          accordionEl.classList.toggle("feedback-accordion--open");
        });
      }

      // One state machine per card — guards double-clicks and enforces valid transitions.
      // The SM instance is closed over by both handlers and GC'd when the list is cleared.
      const sm = new SuggestionStateMachine();

      if (acceptBtn) {
        acceptBtn.addEventListener("click", () =>
          handleAcceptSuggestion(s, li, acceptBtn, rejectBtn, sm)
        );
      }
      if (rejectBtn) {
        rejectBtn.addEventListener("click", () =>
          handleRejectSuggestion(s, li, acceptBtn, rejectBtn, sm)
        );
      }
    }
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

/** Returns the currently selected genre ID from the dropdown. */
function getSelectedGenero(): string {
  const select = document.getElementById("profile-select") as HTMLSelectElement;
  return select.value;
}

// ---------------------------------------------------------------------------
// Accept / Reject — DOM helpers
// ---------------------------------------------------------------------------

/**
 * Appends a small annotation span to a suggestion card `<li>`.
 * Uses `textContent` (never innerHTML) — XSS-safe.
 */
function appendNote(li: HTMLElement, text: string, className: string): void {
  const note = document.createElement("span");
  note.className = className;
  note.textContent = text;
  li.appendChild(note);
}

/**
 * Updates the DOM for a suggestion card based on the SM's terminal state.
 *
 * - `accepted` / `rejected`: removes actions div, adds state class.
 * - `already-resolved`: same + warning note "(ya resuelto)".
 * - `error`: re-enables buttons, shows error in the status bar.
 * - `pending` / `resolving`: no-op (not terminal — should not be called).
 */
function applySuggestionCardState(
  li: HTMLElement,
  state: SuggestionState,
  acceptBtn: HTMLButtonElement | null,
  rejectBtn: HTMLButtonElement | null,
  errorMessage?: string
): void {
  switch (state) {
    case "accepted":
    case "rejected":
      li.querySelector(".result-actions")?.remove();
      li.classList.add(`result-${state}`);
      break;

    case "already-resolved":
      li.querySelector(".result-actions")?.remove();
      li.classList.add("result-already-resolved");
      appendNote(li, "(ya resuelto)", "result-already-resolved-note");
      break;

    case "error":
      if (acceptBtn) acceptBtn.disabled = false;
      if (rejectBtn) rejectBtn.disabled = false;
      showStatus(errorMessage ?? "Error desconocido al resolver sugerencia", "error");
      break;

    default:
      break;
  }
}

// ---------------------------------------------------------------------------
// Accept / Reject Suggestion Handlers
// ---------------------------------------------------------------------------

/**
 * Handles the Accept button click on a suggestion card.
 *
 * Uses a `SuggestionStateMachine` to guard against double-clicks and enforce
 * valid transitions. Sends positive feedback only on explicit acceptance.
 */
async function handleAcceptSuggestion(
  suggestion: Suggestion,
  li: HTMLElement,
  acceptBtn: HTMLButtonElement | null,
  rejectBtn: HTMLButtonElement | null,
  sm: SuggestionStateMachine
): Promise<void> {
  if (!sm.canTransition("resolving")) return;

  sm.transition("resolving");
  if (acceptBtn) acceptBtn.disabled = true;
  if (rejectBtn) rejectBtn.disabled = true;

  const result = await documentPort.acceptSuggestion(suggestion);

  // cc-not-found: terminal UI (remove actions, amber note), SM stays at error
  if (result.status === "cc-not-found") {
    sm.transition("error");
    li.querySelector(".result-actions")?.remove();
    li.classList.add("result-cc-not-found");
    appendNote(li, "(aplicación falló)", "result-cc-not-found-note");
    return;
  }

  const targetState = mapResultStatusToState(result.status);
  sm.transition(targetState);
  applySuggestionCardState(li, sm.state, acceptBtn, rejectBtn, result.error);

  // Positive feedback only on explicit acceptance from the taskpane
  if (sm.state === "accepted") {
    const textarea = li.querySelector(".feedback-textarea") as (HTMLTextAreaElement & { value?: string }) | null;
    const commentText = textarea?.value?.trim();
    const payload: FeedbackPayload = {
      category: suggestion.category,
      originalText: suggestion.originalText,
      ...(suggestion.suggestedText !== undefined ? { suggestedText: suggestion.suggestedText } : {}),
      justification: suggestion.justification,
      rating: "positive",
      severity: suggestion.severity,
      ...(commentText ? { comment: commentText } : {}),
    };
    void feedbackPort.sendFeedback(payload);
  }
}

/**
 * Handles the Reject button click on a suggestion card.
 *
 * Uses a `SuggestionStateMachine` to guard against double-clicks and enforce
 * valid transitions. Sends negative feedback only on explicit rejection.
 */
async function handleRejectSuggestion(
  suggestion: Suggestion,
  li: HTMLElement,
  acceptBtn: HTMLButtonElement | null,
  rejectBtn: HTMLButtonElement | null,
  sm: SuggestionStateMachine
): Promise<void> {
  if (!sm.canTransition("resolving")) return;

  sm.transition("resolving");
  if (acceptBtn) acceptBtn.disabled = true;
  if (rejectBtn) rejectBtn.disabled = true;

  const result = await documentPort.rejectSuggestion(suggestion);

  // cc-not-found: terminal UI (remove actions, amber note), SM stays at error
  if (result.status === "cc-not-found") {
    sm.transition("error");
    li.querySelector(".result-actions")?.remove();
    li.classList.add("result-cc-not-found");
    appendNote(li, "(aplicación falló)", "result-cc-not-found-note");
    return;
  }

  const targetState = mapResultStatusToState(result.status);
  sm.transition(targetState);
  applySuggestionCardState(li, sm.state, acceptBtn, rejectBtn, result.error);

  // Negative feedback only on explicit rejection from the taskpane
  if (sm.state === "rejected") {
    const textarea = li.querySelector(".feedback-textarea") as (HTMLTextAreaElement & { value?: string }) | null;
    const commentText = textarea?.value?.trim();
    const payload: FeedbackPayload = {
      category: suggestion.category,
      originalText: suggestion.originalText,
      ...(suggestion.suggestedText !== undefined ? { suggestedText: suggestion.suggestedText } : {}),
      justification: suggestion.justification,
      rating: "negative",
      severity: suggestion.severity,
      ...(commentText ? { comment: commentText } : {}),
    };
    void feedbackPort.sendFeedback(payload);
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
  // Guard: prevent concurrent runs using the State Machine
  if (stateMachine.isRunning) {
    console.warn("⚠️ [Taskpane] Pipeline ya en ejecución — ignorando click");
    return;
  }

  setAnalyzeLoading(true);
  document.getElementById("results-panel")!.style.display = "none";
  document.getElementById("cleanup-section")!.style.display = "none";

  const emitter = new PipelineEventEmitter();
  const ctx: PipelineContext = {
    documentPort,
    analysisPort,
    emitter,
    genero: getSelectedGenero(),
    maxChunkSize: DEFAULT_MAX_CHUNK_SIZE,
  };

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
      showStatus(reason, (ctx.chunkErrors?.length ?? 0) > 0 ? "error" : "success");
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
