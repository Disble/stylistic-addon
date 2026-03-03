/* global document, Office */

/**
 * Task-pane orchestrator — wires UI events to the Word API and analyzer layers.
 *
 * Responsibilities:
 * - Initialize the task pane when Office.js is ready.
 * - Bind DOM event handlers to the buttons in `taskpane.html`.
 * - Delegate document I/O to {@link wordApi} and text analysis to {@link analyzer}.
 * - Present results, errors, and loading states to the user.
 *
 * This module contains **no** direct Word API calls — all go through `wordApi`.
 * It also contains **no** analysis logic — all goes through `analyzer`.
 *
 * @module taskpane
 */

import { getDocumentText, insertSuggestionsAsTrackedChanges } from "../lib/wordApi";
import { analyze } from "../lib/analyzer";
import { Suggestion, InsertionResult } from "../lib/types";

/** Maximum number of characters displayed in the document preview pane. */
const PREVIEW_MAX_CHARS = 200;

/** Duration (ms) before the status bar message auto-hides. */
const STATUS_DISPLAY_MS = 4000;

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
    document.getElementById("btn-read")!.onclick = handleReadDocument;
    document.getElementById("btn-analyze")!.onclick = handleAnalyze;
  }
});

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------

/**
 * Displays a temporary status message at the bottom of the task pane.
 * The message auto-hides after {@link STATUS_DISPLAY_MS} milliseconds.
 *
 * @param message - The text to display.
 * @param type    - Visual style: `"success"` (green) or `"error"` (red).
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
 * Renders a truncated preview of the document text in the `<pre>` element.
 * Text beyond {@link PREVIEW_MAX_CHARS} is replaced with "...".
 *
 * @param text - The full document text.
 */
function showPreview(text: string): void {
  const preview = document.getElementById("doc-preview")!;
  const truncated = text.length > PREVIEW_MAX_CHARS ? text.substring(0, PREVIEW_MAX_CHARS) + "..." : text;
  preview.textContent = truncated;
  preview.style.display = "block";
}

/**
 * Toggles the "Analizar y sugerir" button between normal and loading states.
 * While loading, the button is disabled and its label reads "Analizando...".
 *
 * @param loading - `true` to enter loading state, `false` to restore.
 */
function setAnalyzeLoading(loading: boolean): void {
  const btn = document.getElementById("btn-analyze") as HTMLButtonElement;
  const label = document.getElementById("btn-analyze-label")!;
  btn.disabled = loading;
  label.textContent = loading ? "Analizando..." : "Analizar y sugerir";
}

/**
 * Renders the results panel showing each suggestion and its outcome.
 *
 * For each suggestion, displays:
 * - **Applied**: original text (struck through) → suggested text (green) + justification.
 * - **Failed**: "Not found" message (red) + justification.
 *
 * @param suggestions - All suggestions produced by the analyzer.
 * @param result      - The {@link InsertionResult} from the Word API layer.
 */
function renderResults(
  suggestions: Suggestion[],
  result: InsertionResult
): void {
  const panel = document.getElementById("results-panel")!;
  const summary = document.getElementById("results-summary")!;
  const list = document.getElementById("results-list")!;

  const total = suggestions.length;
  const applied = result.successCount;
  const failed = result.failedSuggestions.length;

  summary.textContent = `${applied} de ${total} sugerencias aplicadas como Track Changes.`
    + (failed > 0 ? ` ${failed} no encontrada(s) en el texto.` : "");

  list.innerHTML = "";

  for (const s of suggestions) {
    const li = document.createElement("li");
    const isFailed = result.failedSuggestions.some((f) => f.id === s.id);

    if (isFailed) {
      li.innerHTML =
        `<span class="result-failed">No encontrado: "${escapeHtml(s.originalText)}"</span>` +
        `<span class="result-justification">${escapeHtml(s.justification)}</span>`;
    } else {
      li.innerHTML =
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

/**
 * Escapes a string for safe insertion into innerHTML.
 * Uses the browser's own text-node escaping via a temporary `<div>`.
 *
 * @param str - Raw string that may contain HTML-special characters.
 * @returns The HTML-escaped string (safe for innerHTML assignment).
 */
function escapeHtml(str: string): string {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

/**
 * Translates a caught error into a user-friendly message.
 *
 * Recognizes common Office.js error codes and maps them to clear descriptions.
 * Falls back to the raw error message for unknown errors.
 *
 * @param error - The caught error (may be an Error, OfficeExtension.Error, or unknown).
 * @returns A human-readable error description suitable for the status bar.
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

// ---------------------------------------------------------------------------
// Event handlers
// ---------------------------------------------------------------------------

/**
 * Handles the "Leer documento" button click.
 * Reads the full document text and displays a truncated preview.
 */
async function handleReadDocument(): Promise<void> {
  try {
    const text = await getDocumentText();

    if (!text || text.trim().length === 0) {
      showStatus("El documento está vacío.", "error");
      return;
    }

    showPreview(text);
    showStatus(`Documento leído: ${text.length} caracteres en total.`, "success");
  } catch (error) {
    showStatus(toUserMessage(error), "error");
  }
}

/**
 * Handles the "Analizar y sugerir" button click.
 *
 * Full flow:
 * 1. Reads the document text via {@link getDocumentText}.
 * 2. Runs the analyzer to produce {@link Suggestion} objects.
 * 3. Inserts suggestions as tracked changes via {@link insertSuggestionsAsTrackedChanges}.
 * 4. Renders the results panel and shows a status summary.
 *
 * Handles edge cases: empty document, no suggestions found, partial failures,
 * and Office.js errors (translated via {@link toUserMessage}).
 */
async function handleAnalyze(): Promise<void> {
  setAnalyzeLoading(true);

  try {
    const text = await getDocumentText();

    if (!text || text.trim().length === 0) {
      showStatus("El documento está vacío. Escribe algo primero.", "error");
      return;
    }

    const suggestions = analyze(text);

    if (suggestions.length === 0) {
      showStatus("No se encontraron sugerencias editoriales.", "success");
      document.getElementById("results-panel")!.style.display = "none";
      return;
    }

    const result = await insertSuggestionsAsTrackedChanges(suggestions);
    renderResults(suggestions, result);

    if (result.failedSuggestions.length > 0 && result.successCount > 0) {
      showStatus(
        `${result.successCount} aplicada(s), ${result.failedSuggestions.length} no encontrada(s).`,
        "success"
      );
    } else if (result.successCount > 0) {
      showStatus(
        `${result.successCount} sugerencia(s) insertada(s) como Track Changes.`,
        "success"
      );
    } else {
      showStatus("Ninguna sugerencia pudo aplicarse al documento actual.", "error");
    }
  } catch (error) {
    showStatus(toUserMessage(error), "error");
  } finally {
    setAnalyzeLoading(false);
  }
}
