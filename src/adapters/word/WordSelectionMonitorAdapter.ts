/* global Office, Word, console */

import type { SelectionSnapshot } from "../../domain/selection/SelectionSnapshot.types";
import type {
  SelectionListener,
  SelectionMonitorOptions,
} from "./WordSelectionMonitorAdapter.types";
import {
  DEFAULT_PREVIEW_MAX_CHARS,
  EMPTY_SELECTION_SNAPSHOT,
} from "./WordSelectionMonitorAdapter.constants";

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

  private latestSnapshot: SelectionSnapshot = EMPTY_SELECTION_SNAPSHOT;

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
      this.emit(EMPTY_SELECTION_SNAPSHOT);
    }
  }

  private buildSnapshot(rawText: string): SelectionSnapshot {
    const trimmedHasContent = rawText.trim().length > 0;
    if (!trimmedHasContent) {
      return EMPTY_SELECTION_SNAPSHOT;
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
