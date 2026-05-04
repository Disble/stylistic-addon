/* global document, HTMLSelectElement */

import {
  hideTaskpaneProgress,
  setTaskpaneAnalyzeLoading,
  setTaskpaneCleanupVisible,
  setTaskpaneDisableTrackChangesCtaVisible,
  showTaskpaneStatus,
  updateTaskpaneProgress,
} from "./TaskpaneShellStore";

/**
 * Taskpane UI primitives — pure DOM helpers with zero business logic.
 *
 * Every function here operates exclusively on the DOM. No adapter calls,
 * no pipeline state, no mediator. Imported by both `taskpane.ts` and
 * `SuggestionCardRenderer.ts`.
 *
 * @module TaskpaneUi
 */

/** Returns a required DOM element by id or throws with a clear error. */
export function getRequiredElement<T extends HTMLElement = HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing required DOM element: ${id}`);
  }
  return element as T;
}

/**
 * Displays a temporary status message at the bottom of the task pane.
 * Auto-hides after {@link STATUS_DISPLAY_MS} milliseconds.
 */
export function showStatus(message: string, type: "success" | "error"): void {
  showTaskpaneStatus(message, type);
}

/**
 * Toggles the "Analizar y sugerir" button between normal and loading states.
 */
export function setAnalyzeLoading(loading: boolean): void {
  setTaskpaneAnalyzeLoading(loading);
}

/**
 * Updates the progress bar and text in the progress area.
 */
export function updateProgress(current: number, total: number, message: string): void {
  updateTaskpaneProgress(current, total, message);
}

/**
 * Hides the progress bar (called on pipeline completion or abort).
 */
export function hideProgress(): void {
  hideTaskpaneProgress();
}

/**
 * Syncs the cleanup CTA visibility from document-derived workflow semantics.
 */
export function setCleanupCtaVisible(visible: boolean): void {
  setTaskpaneCleanupVisible(visible);
}

/**
 * Syncs the Track Changes CTA visibility from document-derived workflow semantics.
 */
export function setDisableTrackChangesCtaVisible(visible: boolean): void {
  setTaskpaneDisableTrackChangesCtaVisible(visible);
}

/**
 * Translates a caught error into a user-friendly Spanish message.
 * Recognizes common Office.js error codes.
 */
export function toUserMessage(error: unknown): string {
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

/** Returns the currently selected genre ID from the dropdown. */
export function getSelectedGenero(): string {
  const select = document.getElementById("profile-select") as HTMLSelectElement;
  return select.value;
}

/**
 * Appends a small annotation span to a suggestion card element.
 * Uses `textContent` (never innerHTML) — XSS-safe.
 */
export function appendNote(li: HTMLElement, text: string, className: string): void {
  const note = document.createElement("span");
  note.className = className;
  note.textContent = text;
  li.appendChild(note);
}
