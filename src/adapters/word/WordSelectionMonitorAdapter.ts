/* global Office, Word, console */

import type { SelectionSnapshot } from "../../domain/selection/SelectionSnapshot.types";

const DEFAULT_PREVIEW_MAX_CHARS = 80;

const EMPTY_SNAPSHOT: SelectionSnapshot = {
  hasSelection: false,
  charCount: 0,
  preview: "",
};

interface SelectionMonitorOptions {
  previewMaxChars?: number;
}

type SelectionListener = (snapshot: SelectionSnapshot) => void;

/**
 * Bridges Office.js `DocumentSelectionChanged` events into a multiplexed
 * selection-snapshot stream. Subscribers receive an initial snapshot and one
 * per host change. The Office handler is registered lazily on the first
 * subscriber and kept while the adapter exists.
 */
export class WordSelectionMonitorAdapter {
  private readonly listeners = new Set<SelectionListener>();

  private readonly previewMaxChars: number;

  private isHandlerRegistered = false;

  private latestSnapshot: SelectionSnapshot = EMPTY_SNAPSHOT;

  constructor(options: SelectionMonitorOptions = {}) {
    this.previewMaxChars = options.previewMaxChars ?? DEFAULT_PREVIEW_MAX_CHARS;
  }

  /** Subscribes to selection changes and returns an unsubscribe function. */
  subscribe(listener: SelectionListener): () => void {
    this.listeners.add(listener);
    this.ensureHandlerRegistered();
    void this.refreshSnapshot();
    return () => {
      this.listeners.delete(listener);
    };
  }

  private ensureHandlerRegistered(): void {
    if (this.isHandlerRegistered) {
      return;
    }
    this.isHandlerRegistered = true;
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => {
        void this.refreshSnapshot();
      },
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn(
            "⚠️ [WordSelectionMonitor] No se pudo registrar el handler de selección.",
            asyncResult
          );
        }
      }
    );
  }

  private async refreshSnapshot(): Promise<void> {
    try {
      const selectionText = await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        return selection.text ?? "";
      });
      this.emit(this.buildSnapshot(selectionText));
    } catch (error) {
      console.warn("⚠️ [WordSelectionMonitor] Error leyendo selección:", error);
      this.emit(EMPTY_SNAPSHOT);
    }
  }

  private buildSnapshot(rawText: string): SelectionSnapshot {
    const trimmedHasContent = rawText.trim().length > 0;
    if (!trimmedHasContent) {
      return EMPTY_SNAPSHOT;
    }
    const normalized = rawText.replace(/\s+/g, " ").trim();
    const preview =
      normalized.length > this.previewMaxChars
        ? `${normalized.slice(0, this.previewMaxChars)}…`
        : normalized;
    return {
      hasSelection: true,
      charCount: rawText.length,
      preview,
    };
  }

  private emit(snapshot: SelectionSnapshot): void {
    this.latestSnapshot = snapshot;
    for (const listener of this.listeners) {
      listener(snapshot);
    }
  }
}
