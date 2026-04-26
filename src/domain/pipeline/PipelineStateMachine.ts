/* global console */

/**
 * Pipeline State Machine — enforces valid state transitions for the analysis
 * pipeline lifecycle.
 *
 * **State pattern** implementation. All valid transitions are declared
 * explicitly. Any attempt to transition to an invalid state throws, making
 * bugs immediately visible rather than silently corrupting pipeline state.
 *
 * States:
 * - `idle`       → Pipeline is not running. Ready to start.
 * - `reading`    → Reading text from document or selection.
 * - `connecting` → Verifying backend connectivity.
 * - `chunking`   → Splitting text at paragraph boundaries.
 * - `analyzing`  → Sending chunks to the Mastra workflow sequentially.
 * - `applying`   → Applying suggestions as tracked changes in Word.
 * - `done`       → Pipeline completed. Results available.
 * - `error`      → Pipeline aborted. Error message available.
 *
 * @module PipelineStateMachine
 */

import type { PipelineState } from "./PipelineStateMachine.types";

export type { PipelineState } from "./PipelineStateMachine.types";

const TRANSITIONS: Record<PipelineState, PipelineState[]> = {
  idle: ["reading"],
  reading: ["connecting", "idle", "error"],
  connecting: ["chunking", "idle", "error"],
  chunking: ["analyzing", "idle", "error"],
  analyzing: ["applying", "idle", "error"],
  applying: ["done", "idle", "error"],
  done: ["idle"],
  error: ["idle"],
};

/**
 * Manages the lifecycle of a single pipeline run.
 *
 * Usage:
 * ```typescript
 * const sm = new PipelineStateMachine();
 * sm.transition("reading");   // idle → reading
 * sm.transition("connecting"); // reading → connecting
 * // ...
 * sm.transition("done");
 * sm.reset();                 // done → idle (ready for next run)
 * ```
 */
export class PipelineStateMachine {
  private current: PipelineState = "idle";

  /** The current state. */
  get state(): PipelineState {
    return this.current;
  }

  /** Returns `true` if the pipeline is currently executing (not idle/done/error). */
  get isRunning(): boolean {
    return (
      this.current !== "idle" &&
      this.current !== "done" &&
      this.current !== "error"
    );
  }

  /**
   * Returns `true` if the given transition is valid from the current state.
   *
   * @param to - Target state.
   */
  canTransition(to: PipelineState): boolean {
    return TRANSITIONS[this.current].includes(to);
  }

  /**
   * Transitions to the given state.
   *
   * @param to - Target state.
   * @throws {Error} If the transition is not valid from the current state.
   */
  transition(to: PipelineState): void {
    if (!this.canTransition(to)) {
      throw new Error(
        `[PipelineStateMachine] Invalid transition: "${this.current}" → "${to}". ` +
          `Allowed: [${TRANSITIONS[this.current].join(", ")}]`,
      );
    }
    console.log(`🔄 [StateMachine] ${this.current} → ${to}`);
    this.current = to;
  }

  /**
   * Resets the state machine to `idle`.
   * Idempotent convenience for returning to `idle` — safe to call in `finally` blocks.
   */
  reset(): void {
    console.log(`🔄 [StateMachine] reset → idle (was: ${this.current})`);
    this.current = "idle";
  }
}
