/**
 * Pipeline Event Emitter — Observer pattern for decoupled progress reporting.
 *
 * Allows multiple consumers (UI progress bar, logging, analytics) to react
 * to pipeline events without the pipeline needing to know about them.
 *
 * The `taskpane.ts` UI registers an observer at startup. Future observers
 * (telemetry, structured logging) can be added without touching the pipeline.
 *
 * @module PipelineEvents
 */

import type { ApplySuggestionsResult } from "../DocumentApplication.types";
import type { Suggestion } from "../suggestion/Suggestion.types";
import type { PipelineState } from "./PipelineStateMachine.types";

// ---------------------------------------------------------------------------
// Observer interface
// ---------------------------------------------------------------------------

/**
 * Observer for pipeline lifecycle events.
 * All methods are optional — implement only what you need.
 */
export interface PipelineObserver {
  /** Called when a pipeline phase begins. */
  onPhaseStart?(phase: PipelineState, message: string): void;

  /** Called periodically during a phase to report incremental progress. */
  onProgress?(current: number, total: number, message: string): void;

  /** Called when a pipeline phase completes successfully. */
  onPhaseComplete?(phase: PipelineState): void;

  /** Called when the pipeline encounters a recoverable or fatal error. */
  onError?(phase: PipelineState, error: Error | string): void;

  /** Called when the pipeline completes (success or partial success). */
  onComplete?(
    suggestions: Suggestion[],
    result: ApplySuggestionsResult,
    chunkErrors: string[],
    isSelection: boolean,
  ): void;

  /** Called when the pipeline is aborted before reaching `done`. */
  onAbort?(reason: string): void;
}

// ---------------------------------------------------------------------------
// Event Emitter
// ---------------------------------------------------------------------------

/**
 * Manages a list of `PipelineObserver` instances and dispatches events to all.
 *
 * Usage:
 * ```typescript
 * const emitter = new PipelineEventEmitter();
 * emitter.subscribe({ onProgress: (c, t, msg) => updateUI(c, t, msg) });
 * emitter.emitProgress(1, 5, "Analizando fragmento 1 de 5...");
 * ```
 */
export class PipelineEventEmitter {
  private observers: PipelineObserver[] = [];

  private notifyObservers(notify: (observer: PipelineObserver) => void): void {
    for (const observer of this.observers) {
      try {
        notify(observer);
      } catch {
        // Ignore observer failures so other observers still receive the event.
      }
    }
  }

  /**
   * Registers an observer. The observer will receive all subsequent events
   * until `unsubscribe()` is called.
   */
  subscribe(observer: PipelineObserver): void {
    this.observers.push(observer);
  }

  /**
   * Removes a previously registered observer.
   */
  unsubscribe(observer: PipelineObserver): void {
    this.observers = this.observers.filter((o) => o !== observer);
  }

  /** Removes all registered observers. */
  clear(): void {
    this.observers = [];
  }

  emitPhaseStart(phase: PipelineState, message: string): void {
    this.notifyObservers((observer) => observer.onPhaseStart?.(phase, message));
  }

  emitProgress(current: number, total: number, message: string): void {
    this.notifyObservers((observer) =>
      observer.onProgress?.(current, total, message),
    );
  }

  emitPhaseComplete(phase: PipelineState): void {
    this.notifyObservers((observer) => observer.onPhaseComplete?.(phase));
  }

  emitError(phase: PipelineState, error: Error | string): void {
    this.notifyObservers((observer) => observer.onError?.(phase, error));
  }

  emitComplete(
    suggestions: Suggestion[],
    result: ApplySuggestionsResult,
    chunkErrors: string[],
    isSelection: boolean,
  ): void {
    this.notifyObservers((observer) =>
      observer.onComplete?.(suggestions, result, chunkErrors, isSelection),
    );
  }

  emitAbort(reason: string): void {
    this.notifyObservers((observer) => observer.onAbort?.(reason));
  }
}
