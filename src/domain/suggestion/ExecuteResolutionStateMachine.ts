/* global console */

import type { ResolutionPhase } from "../types";

export type ExecuteResolutionState =
  | "idle"
  | "locating"
  | "observing-before"
  | "executing"
  | "cleaning-comment"
  | "cleaning-anchor"
  | "inspecting-after"
  | "completed"
  | "failed";

const TRANSITIONS: Record<ExecuteResolutionState, ExecuteResolutionState[]> = {
  idle: ["locating", "failed"],
  locating: ["observing-before", "failed", "completed"],
  "observing-before": ["executing", "failed", "completed"],
  executing: ["cleaning-comment", "failed"],
  "cleaning-comment": ["cleaning-anchor", "failed"],
  "cleaning-anchor": ["inspecting-after", "failed"],
  "inspecting-after": ["completed", "failed"],
  completed: [],
  failed: [],
};

export class InvalidExecuteResolutionTransitionError extends Error {
  constructor(
    from: ExecuteResolutionState,
    to: ExecuteResolutionState,
    allowed: ExecuteResolutionState[],
  ) {
    super(
      `[ExecuteResolutionStateMachine] Invalid transition: "${from}" → "${to}". ` +
        `Allowed: [${allowed.join(", ")}]`,
    );
    this.name = "InvalidExecuteResolutionTransitionError";
  }
}

/** Owns internal execute-phase transitions for one suggestion resolution workflow. */
export class ExecuteResolutionStateMachine {
  private current: ExecuteResolutionState = "idle";
  private lastActivePhase: ResolutionPhase | null = null;

  /** Current internal resolution state. */
  get state(): ExecuteResolutionState {
    return this.current;
  }

  /** Returns true when no more transitions are possible. */
  get isTerminal(): boolean {
    return TRANSITIONS[this.current].length === 0;
  }

  /** Returns the semantic phase represented by the current state. */
  get currentPhase(): ResolutionPhase | null {
    if (this.current === "failed") {
      return this.lastActivePhase;
    }

    switch (this.current) {
      case "locating":
        return "locate";
      case "observing-before":
        return "observe-before";
      case "executing":
        return "execute";
      case "cleaning-comment":
        return "cleanup-comment";
      case "cleaning-anchor":
        return "cleanup-anchor";
      case "inspecting-after":
        return "inspect-after";
      default:
        return null;
    }
  }

  /** Returns true when the given transition is allowed from the current state. */
  canTransition(to: ExecuteResolutionState): boolean {
    return TRANSITIONS[this.current].includes(to);
  }

  /** Applies one validated transition. */
  transition(to: ExecuteResolutionState): void {
    if (!this.canTransition(to)) {
      throw new InvalidExecuteResolutionTransitionError(
        this.current,
        to,
        TRANSITIONS[this.current],
      );
    }

    console.log(`🔄 [ExecuteResolutionStateMachine] ${this.current} → ${to}`);
    if (to !== "completed" && to !== "failed") {
      this.lastActivePhase = this.currentPhaseForState(to);
    }
    this.current = to;
  }

  /** Marks the workflow as failed while preserving the last active phase. */
  fail(): void {
    this.transition("failed");
  }

  /** Resolves the semantic phase for one non-terminal active state. */
  private currentPhaseForState(
    state: Exclude<ExecuteResolutionState, "idle" | "completed" | "failed">,
  ): ResolutionPhase {
    switch (state) {
      case "locating":
        return "locate";
      case "observing-before":
        return "observe-before";
      case "executing":
        return "execute";
      case "cleaning-comment":
        return "cleanup-comment";
      case "cleaning-anchor":
        return "cleanup-anchor";
      case "inspecting-after":
        return "inspect-after";
    }
  }
}
