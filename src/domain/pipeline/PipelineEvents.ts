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

import { PipelineState, InsertionResult, Suggestion } from "../types";

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
    result: InsertionResult,
    chunkErrors: string[],
    isSelection: boolean
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
    for (const o of this.observers) o.onPhaseStart?.(phase, message);
  }

  emitProgress(current: number, total: number, message: string): void {
    for (const o of this.observers) o.onProgress?.(current, total, message);
  }

  emitPhaseComplete(phase: PipelineState): void {
    for (const o of this.observers) o.onPhaseComplete?.(phase);
  }

  emitError(phase: PipelineState, error: Error | string): void {
    for (const o of this.observers) o.onError?.(phase, error);
  }

  emitComplete(
    suggestions: Suggestion[],
    result: InsertionResult,
    chunkErrors: string[],
    isSelection: boolean
  ): void {
    for (const o of this.observers) o.onComplete?.(suggestions, result, chunkErrors, isSelection);
  }

  emitAbort(reason: string): void {
    for (const o of this.observers) o.onAbort?.(reason);
  }
}
