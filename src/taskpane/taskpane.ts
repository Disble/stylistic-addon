/* global document, Office */

/**
 * Task-pane orchestrator — wires UI events to the Word API and Mastra
 * backend layers.
 *
 * Responsibilities:
 * - Initialize the task pane when Office.js is ready.
 * - Bind DOM event handlers to the UI controls in `taskpane.html`.
 * - Coordinate the multi-phase analysis pipeline:
 *   1. Read document text ({@link wordApi}).
 *   2. Check backend connectivity ({@link mastraClient}).
 *   3. Chunk text ({@link chunker}).
 *   4. Analyze each chunk via the Mastra workflow ({@link mastraClient}).
 *   5. Deduplicate suggestions.
 *   6. Apply tracked changes in batches ({@link wordApi}).
 *   7. Render results and status.
 *
 * This module contains **no** direct Word API calls — all go through `wordApi`.
 * It also contains **no** backend communication — all goes through `mastraClient`.
 * It contains **no** business logic — analysis is entirely server-side.
 *
 * @module taskpane
 */

import {
  getTextToAnalyze,
  applySuggestionsInBatches,
  cleanupResolvedComments,
} from "../lib/wordApi";
import { checkConnection, analyzeChunk } from "../lib/mastraClient";
import { splitText } from "../lib/chunker";
import { DEFAULT_MAX_CHUNK_SIZE } from "../lib/config";
import { Suggestion, InsertionResult, ChunkResult, AnalysisPhase } from "../lib/types";

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
    document.getElementById("btn-analyze")!.onclick = handleAnalyze;
    document.getElementById("btn-cleanup")!.onclick = handleCleanup;
  }
});

// ---------------------------------------------------------------------------
// UI Helpers
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
 * Toggles the "Analizar y sugerir" button between normal and loading states.
 *
 * @param loading - `true` to enter loading state, `false` to restore.
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
 *
 * @param phase   - Current analysis phase.
 * @param current - Current step within the phase (1-based).
 * @param total   - Total steps in the phase.
 * @param message - Human-readable status text.
 */
function updateProgress(
  phase: AnalysisPhase,
  current: number,
  total: number,
  message: string
): void {
  const container = document.getElementById("progress-container")!;
  const bar = document.getElementById("progress-bar")!;
  const text = document.getElementById("progress-text")!;

  container.style.display = "block";

  const percentage = total > 0 ? Math.round((current / total) * 100) : 0;
  bar.style.width = `${percentage}%`;
  text.textContent = message;

  if (phase === "done") {
    setTimeout(() => {
      container.style.display = "none";
    }, 1000);
  }
}

/**
 * Renders the results panel showing each suggestion and its outcome.
 *
 * @param suggestions  - All suggestions produced by the workflow.
 * @param result       - The {@link InsertionResult} from the Word API layer.
 * @param chunkErrors  - Error messages from failed chunks (if any).
 * @param isSelection  - Whether the analysis was scoped to a text selection.
 */
function renderResults(
  suggestions: Suggestion[],
  result: InsertionResult,
  chunkErrors: string[],
  isSelection: boolean = false
): void {
  const panel = document.getElementById("results-panel")!;
  const summary = document.getElementById("results-summary")!;
  const list = document.getElementById("results-list")!;

  const total = suggestions.length;
  const applied = result.successCount;
  const failed = result.failedSuggestions.length;

  const scopePrefix = isSelection ? "Sobre selección — " : "";
  let summaryText = `${scopePrefix}${applied} de ${total} sugerencias aplicadas como Track Changes.`;
  if (failed > 0) {
    summaryText += ` ${failed} no encontrada(s) en el texto.`;
  }
  if (chunkErrors.length > 0) {
    summaryText += ` ${chunkErrors.length} fragmento(s) con error.`;
  }
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

/**
 * Escapes a string for safe insertion into innerHTML.
 *
 * @param str - Raw string that may contain HTML-special characters.
 * @returns The HTML-escaped string.
 */
function escapeHtml(str: string): string {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

/**
 * Translates a caught error into a user-friendly message.
 * Recognizes common Office.js error codes and maps them to descriptions.
 *
 * @param error - The caught error (may be an Error, OfficeExtension.Error, or unknown).
 * @returns A human-readable error description.
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

/**
 * Returns the currently selected profile ID from the dropdown.
 *
 * @returns The selected profile identifier string.
 */
function getSelectedProfile(): string {
  const select = document.getElementById("profile-select") as HTMLSelectElement;
  return select.value;
}

/**
 * Removes duplicate suggestions that target the same original text.
 * When multiple chunks return suggestions for the same phrase, only the first
 * is kept. Comparison is case-insensitive.
 *
 * @param suggestions - The raw suggestion list (may contain duplicates).
 * @returns A filtered list with unique `originalText` values.
 */
function deduplicateByOriginalText(suggestions: Suggestion[]): Suggestion[] {
  const seen = new Set<string>();
  return suggestions.filter((s) => {
    const key = s.originalText.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// ---------------------------------------------------------------------------
// Main Event Handler
// ---------------------------------------------------------------------------

/**
 * Handles the "Analizar y sugerir" button click.
 *
 * Orchestrates the full multi-phase analysis pipeline:
 * 1. Read document text.
 * 2. Validate backend connectivity.
 * 3. Chunk the text for the workflow.
 * 4. Analyze each chunk via the Mastra editorial workflow.
 * 5. Deduplicate and apply suggestions as tracked changes.
 * 6. Render results and status.
 *
 * Handles partial failures gracefully — if some chunks fail, the successful
 * suggestions are still applied. Only a completely empty result set triggers
 * an error state.
 */
async function handleAnalyze(): Promise<void> {
  setAnalyzeLoading(true);
  document.getElementById("results-panel")!.style.display = "none";
  document.getElementById("cleanup-section")!.style.display = "none";

  try {
    // Phase 1: Resolve text source (selection if active, full document otherwise)
    console.log("🚀 [Taskpane] Pipeline iniciado");
    updateProgress("reading", 0, 1, "Leyendo texto...");
    console.log("📖 [Taskpane] Fase 1: Resolviendo texto a analizar...");
    const { text, isSelection } = await getTextToAnalyze();
    const scope = isSelection ? "selección" : "documento";
    console.log(`📖 [Taskpane] ${isSelection ? "Selección leída" : "Documento leído"} — ${text.length} caracteres`);

    if (!text || text.trim().length === 0) {
      console.warn("⚠️ [Taskpane] Texto vacío, abortando");
      showStatus("El documento está vacío. Escribe algo primero.", "error");
      return;
    }

    // Phase 2: Check backend
    console.log("🔌 [Taskpane] Fase 2: Verificando conexión con backend...");
    updateProgress("connecting", 0, 1, "Conectando con el servidor...");
    const connected = await checkConnection();
    console.log(`🔌 [Taskpane] Conexión: ${connected ? "✅ OK" : "❌ FALLO"}`);

    if (!connected) {
      console.error("❌ [Taskpane] Backend no disponible, abortando");
      showStatus(
        "Backend no disponible. Verifica que el servidor Mastra esté ejecutándose.",
        "error"
      );
      return;
    }

    // Phase 3: Chunk text
    console.log(`✂️ [Taskpane] Fase 3: Dividiendo ${scope} en chunks...`);
    const chunks = splitText(text, DEFAULT_MAX_CHUNK_SIZE);
    const profile = getSelectedProfile();
    console.log(`✂️ [Taskpane] ${chunks.length} chunk(s) generados — perfil: "${profile}", ámbito: ${scope}`);

    // Phase 4: Analyze each chunk
    console.log("🤖 [Taskpane] Fase 4: Analizando chunks con Mastra...");
    const allSuggestions: Suggestion[] = [];
    const chunkErrors: string[] = [];

    for (const chunk of chunks) {
      updateProgress(
        "analyzing",
        chunk.index + 1,
        chunks.length,
        `Analizando fragmento ${chunk.index + 1} de ${chunks.length} (${scope})...`
      );

      console.log(
        `🤖 [Taskpane] Enviando chunk ${chunk.index + 1}/${chunks.length} (${chunk.text.length} chars)`
      );
      const chunkResult: ChunkResult = await analyzeChunk(chunk, profile, "es");
      console.log(
        `🤖 [Taskpane] Chunk ${chunk.index + 1} → ${chunkResult.suggestions.length} sugerencia(s)${chunkResult.error ? " ⚠️ con error: " + chunkResult.error : ""}`
      );

      allSuggestions.push(...chunkResult.suggestions);
      if (chunkResult.error) {
        chunkErrors.push(chunkResult.error);
      }
    }

    // Phase 5: Deduplicate
    console.log(
      `🧹 [Taskpane] Fase 5: Deduplicando — ${allSuggestions.length} sugerencias totales`
    );
    const uniqueSuggestions = deduplicateByOriginalText(allSuggestions);
    console.log(
      `🧹 [Taskpane] ${uniqueSuggestions.length} sugerencias únicas (${allSuggestions.length - uniqueSuggestions.length} duplicadas removidas)`
    );

    if (uniqueSuggestions.length === 0) {
      console.warn(`⚠️ [Taskpane] Sin sugerencias. Errores de chunks: ${chunkErrors.length}`);
      if (chunkErrors.length > 0) {
        showStatus(
          `El análisis falló en ${chunkErrors.length} fragmento(s). Intenta de nuevo.`,
          "error"
        );
      } else {
        showStatus("No se encontraron sugerencias editoriales.", "success");
      }
      updateProgress("done", 1, 1, "");
      return;
    }

    // Phase 6: Apply as tracked changes
    console.log(
      `📝 [Taskpane] Fase 6: Aplicando ${uniqueSuggestions.length} sugerencias como Track Changes...`
    );
    const result = await applySuggestionsInBatches(uniqueSuggestions, updateProgress);
    console.log(
      `📝 [Taskpane] Resultado: ${result.successCount} aplicadas, ${result.failedSuggestions.length} fallidas`
    );

    // Phase 7: Render results
    console.log("🎨 [Taskpane] Fase 7: Renderizando resultados");
    updateProgress("done", 1, 1, "");
    renderResults(uniqueSuggestions, result, chunkErrors, isSelection);
    console.log("✅ [Taskpane] Pipeline completado exitosamente");

    // Show cleanup button if any suggestions were applied
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
  } catch (error) {
    console.error("💥 [Taskpane] Error en pipeline:", error);
    updateProgress("done", 1, 1, "");
    showStatus(toUserMessage(error), "error");
  } finally {
    setAnalyzeLoading(false);
  }
}

// ---------------------------------------------------------------------------
// Comment Cleanup Handler
// ---------------------------------------------------------------------------

/**
 * Handles the "Limpiar comentarios resueltos" button click.
 * Deletes Stylistic comments whose tracked changes have been resolved.
 */
async function handleCleanup(): Promise<void> {
  console.log("🧽 [Taskpane] Iniciando limpieza de comentarios resueltos...");
  const btn = document.getElementById("btn-cleanup") as HTMLButtonElement;
  const label = document.getElementById("btn-cleanup-label")!;

  btn.disabled = true;
  label.textContent = "Limpiando...";

  try {
    const { deleted, kept } = await cleanupResolvedComments();
    console.log(`🧽 [Taskpane] Limpieza: ${deleted} eliminados, ${kept} conservados`);

    showStatus(`${deleted} comentario(s) eliminado(s), ${kept} conservado(s).`, "success");

    // If nothing left to clean, hide the button
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
