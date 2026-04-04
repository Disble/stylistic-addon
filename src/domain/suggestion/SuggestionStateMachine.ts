/* global console */

/**
 * Suggestion Card State Machine — enforces valid state transitions for the
 * accept/reject lifecycle of a single suggestion card in the taskpane.
 *
 * Follows the `PipelineStateMachine` pattern exactly: explicit transition table,
 * throws on invalid transitions, logs every valid transition.
 *
 * States:
 * - `pending`          → Card awaiting user action. Buttons enabled.
 * - `resolving`        → Async Word API call in-flight. Buttons disabled.
 * - `accepted`         → User accepted from taskpane. Terminal.
 * - `rejected`         → User rejected from taskpane. Terminal.
 * - `already-resolved` → CC found but TCs already gone (resolved via Review pane). Terminal.
 * - `unobservable`     → Word did not expose enough evidence to confirm resolution. Non-terminal.
 * - `identity-lost`    → Word exposed corrupt/incomplete compound identity metadata. Terminal warning.
 * - `error`            → Word API threw. Non-terminal — user may retry.
 *
 * @module SuggestionStateMachine
 */

import type { SuggestionActionResult, SuggestionState } from "../types";

const TRANSITIONS: Record<SuggestionState, SuggestionState[]> = {
  pending: ["resolving"],
  resolving: [
    "accepted",
    "rejected",
    "already-resolved",
    "unobservable",
    "identity-lost",
    "error",
  ],
  accepted: [],
  rejected: [],
  "already-resolved": [],
  unobservable: ["resolving"],
  "identity-lost": [],
  error: ["resolving"],
};

export class InvalidSuggestionTransitionError extends Error {
  constructor(
    from: SuggestionState,
    to: SuggestionState,
    allowed: SuggestionState[],
  ) {
    super(
      `[SuggestionStateMachine] Invalid transition: "${from}" → "${to}". ` +
        `Allowed: [${allowed.join(", ")}]`,
    );
    this.name = "InvalidSuggestionTransitionError";
  }
}

export class SuggestionStateMachine {
  private current: SuggestionState = "pending";

  /** The current state. */
  get state(): SuggestionState {
    return this.current;
  }

  /** `true` when state is terminal (no further transitions possible). */
  get isTerminal(): boolean {
    return TRANSITIONS[this.current].length === 0;
  }

  /** `true` when an async resolution is in-flight. */
  get isResolving(): boolean {
    return this.current === "resolving";
  }

  /** Returns `true` if the given transition is valid from the current state. */
  canTransition(to: SuggestionState): boolean {
    return TRANSITIONS[this.current].includes(to);
  }

  /**
   * Transitions to the given state.
   * @throws {InvalidSuggestionTransitionError} If the transition is not valid.
   */
  transition(to: SuggestionState): void {
    if (!this.canTransition(to)) {
      throw new InvalidSuggestionTransitionError(
        this.current,
        to,
        TRANSITIONS[this.current],
      );
    }
    console.log(`🔄 [SuggestionStateMachine] ${this.current} → ${to}`);
    this.current = to;
  }

  /**
   * Resets the state machine to `pending`.
   * Idempotent — safe to call from any state.
   */
  reset(): void {
    console.log(
      `🔄 [SuggestionStateMachine] reset → pending (was: ${this.current})`,
    );
    this.current = "pending";
  }
}

/**
 * Maps a `SuggestionActionResult.status` to the corresponding `SuggestionState`.
 *
 * - `"cc-not-found"` and `"not-found"` both map to `"error"` (retryable).
 *   The taskpane distinguishes `"cc-not-found"` visually before calling this function.
 */
export function mapResultStatusToState(
  status: SuggestionActionResult["status"],
): SuggestionState {
  switch (status) {
    case "accepted":
      return "accepted";
    case "rejected":
      return "rejected";
    case "already-resolved":
      return "already-resolved";
    case "unobservable":
      return "unobservable";
    case "identity-lost":
      return "identity-lost";
    case "cc-not-found":
      return "error";
    case "not-found":
      return "error";
    case "error":
      return "error";
  }
}
