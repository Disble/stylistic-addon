/* global console */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { SuggestionState } from "../types";
import {
  InvalidSuggestionTransitionError,
  mapResultStatusToState,
  SuggestionStateMachine,
} from "./SuggestionStateMachine";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const TERMINAL_STATES: SuggestionState[] = [
  "accepted",
  "rejected",
  "already-resolved",
];

/** Returns an SM already transitioned to the given state. */
function machineAt(target: SuggestionState): SuggestionStateMachine {
  const sm = new SuggestionStateMachine();
  const paths: Partial<Record<SuggestionState, SuggestionState[]>> = {
    pending: [],
    resolving: ["resolving"],
    accepted: ["resolving", "accepted"],
    rejected: ["resolving", "rejected"],
    "already-resolved": ["resolving", "already-resolved"],
    error: ["resolving", "error"],
  };
  for (const step of paths[target] ?? []) sm.transition(step);
  return sm;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("SuggestionStateMachine", () => {
  let logSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    logSpy = vi.spyOn(console, "log").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  // -------------------------------------------------------------------------
  // 1. Initial state
  // -------------------------------------------------------------------------

  describe("initial state", () => {
    it("starts in 'pending'", () => {
      const sm = new SuggestionStateMachine();
      expect(sm.state).toBe("pending");
    });

    it("isTerminal is false in 'pending'", () => {
      expect(new SuggestionStateMachine().isTerminal).toBe(false);
    });

    it("canTransition('resolving') is true from 'pending'", () => {
      expect(new SuggestionStateMachine().canTransition("resolving")).toBe(
        true,
      );
    });
  });

  // -------------------------------------------------------------------------
  // 2. Valid transitions — happy paths
  // -------------------------------------------------------------------------

  describe("valid transitions", () => {
    it("pending → resolving → accepted", () => {
      const sm = new SuggestionStateMachine();
      sm.transition("resolving");
      expect(sm.state).toBe("resolving");
      sm.transition("accepted");
      expect(sm.state).toBe("accepted");
      expect(sm.isTerminal).toBe(true);
    });

    it("pending → resolving → rejected", () => {
      const sm = new SuggestionStateMachine();
      sm.transition("resolving");
      sm.transition("rejected");
      expect(sm.state).toBe("rejected");
      expect(sm.isTerminal).toBe(true);
    });

    it("pending → resolving → already-resolved", () => {
      const sm = new SuggestionStateMachine();
      sm.transition("resolving");
      sm.transition("already-resolved");
      expect(sm.state).toBe("already-resolved");
      expect(sm.isTerminal).toBe(true);
    });

    it("pending → resolving → error (non-terminal, retry allowed)", () => {
      const sm = new SuggestionStateMachine();
      sm.transition("resolving");
      sm.transition("error");
      expect(sm.state).toBe("error");
      expect(sm.isTerminal).toBe(false);
      expect(sm.canTransition("resolving")).toBe(true);
    });

    it("error → resolving → accepted (retry cycle)", () => {
      const sm = machineAt("error");
      sm.transition("resolving");
      sm.transition("accepted");
      expect(sm.state).toBe("accepted");
    });

    it("error → resolving → rejected (retry cycle)", () => {
      const sm = machineAt("error");
      sm.transition("resolving");
      sm.transition("rejected");
      expect(sm.state).toBe("rejected");
    });
  });

  // -------------------------------------------------------------------------
  // 3. isTerminal
  // -------------------------------------------------------------------------

  describe("isTerminal", () => {
    it.each(TERMINAL_STATES)("isTerminal is true for '%s'", (state) => {
      expect(machineAt(state).isTerminal).toBe(true);
    });

    it("isTerminal is false for 'resolving'", () => {
      expect(machineAt("resolving").isTerminal).toBe(false);
    });

    it("isTerminal is false for 'error'", () => {
      expect(machineAt("error").isTerminal).toBe(false);
    });
  });

  // -------------------------------------------------------------------------
  // 4. isResolving
  // -------------------------------------------------------------------------

  describe("isResolving", () => {
    it("isResolving is true only when state is 'resolving'", () => {
      const sm = new SuggestionStateMachine();
      expect(sm.isResolving).toBe(false);
      sm.transition("resolving");
      expect(sm.isResolving).toBe(true);
      sm.transition("accepted");
      expect(sm.isResolving).toBe(false);
    });
  });

  // -------------------------------------------------------------------------
  // 5. canTransition — does NOT mutate state
  // -------------------------------------------------------------------------

  describe("canTransition", () => {
    it("does not mutate state", () => {
      const sm = new SuggestionStateMachine();
      sm.canTransition("resolving");
      expect(sm.state).toBe("pending");
    });

    it("returns false for invalid transitions", () => {
      expect(new SuggestionStateMachine().canTransition("accepted")).toBe(
        false,
      );
      expect(new SuggestionStateMachine().canTransition("rejected")).toBe(
        false,
      );
      expect(new SuggestionStateMachine().canTransition("error")).toBe(false);
    });
  });

  // -------------------------------------------------------------------------
  // 6. Invalid transitions
  // -------------------------------------------------------------------------

  describe("invalid transitions", () => {
    it("throws when skipping resolving (pending → accepted)", () => {
      const sm = new SuggestionStateMachine();
      expect(() => sm.transition("accepted")).toThrow(
        InvalidSuggestionTransitionError,
      );
      expect(sm.state).toBe("pending");
    });

    it("throws when skipping resolving (pending → error)", () => {
      const sm = new SuggestionStateMachine();
      expect(() => sm.transition("error")).toThrow(
        InvalidSuggestionTransitionError,
      );
      expect(sm.state).toBe("pending");
    });

    it.each(
      TERMINAL_STATES,
    )("throws from terminal state '%s' → resolving", (state) => {
      const sm = machineAt(state);
      expect(() => sm.transition("resolving")).toThrow(
        InvalidSuggestionTransitionError,
      );
      expect(sm.state).toBe(state);
    });

    it("double-click guard: resolving → resolving throws", () => {
      const sm = machineAt("resolving");
      expect(() => sm.transition("resolving")).toThrow(
        InvalidSuggestionTransitionError,
      );
      expect(sm.state).toBe("resolving");
    });

    it("throws from 'resolving' → 'pending' (no going back)", () => {
      const sm = machineAt("resolving");
      expect(() => sm.transition("pending")).toThrow(
        InvalidSuggestionTransitionError,
      );
    });
  });

  // -------------------------------------------------------------------------
  // 7. InvalidSuggestionTransitionError message format
  // -------------------------------------------------------------------------

  describe("error message", () => {
    it("includes from state, to state, and allowed targets", () => {
      const sm = new SuggestionStateMachine();
      let thrown: Error | undefined;
      try {
        sm.transition("accepted");
      } catch (e) {
        thrown = e as Error;
      }
      expect(thrown).toBeInstanceOf(InvalidSuggestionTransitionError);
      expect(thrown?.message).toContain('"pending"');
      expect(thrown?.message).toContain('"accepted"');
      expect(thrown?.message).toContain("Allowed:");
      expect(thrown?.message).toMatch(/Invalid transition/);
    });

    it("error name is 'InvalidSuggestionTransitionError'", () => {
      let thrown: Error | undefined;
      try {
        new SuggestionStateMachine().transition("accepted");
      } catch (e) {
        thrown = e as Error;
      }
      expect(thrown?.name).toBe("InvalidSuggestionTransitionError");
    });
  });

  // -------------------------------------------------------------------------
  // 8. console.log behavior
  // -------------------------------------------------------------------------

  describe("console.log", () => {
    it("logs on valid transition with [SuggestionStateMachine] prefix", () => {
      const sm = new SuggestionStateMachine();
      sm.transition("resolving");
      expect(logSpy).toHaveBeenCalledWith(
        expect.stringContaining("[SuggestionStateMachine]"),
      );
      expect(logSpy).toHaveBeenCalledWith(
        expect.stringContaining("pending → resolving"),
      );
    });

    it("does NOT log on invalid transition", () => {
      const sm = new SuggestionStateMachine();
      try {
        sm.transition("accepted");
      } catch {
        /* expected */
      }
      expect(logSpy).not.toHaveBeenCalled();
    });

    it("logs on reset with 'was:' context", () => {
      const sm = machineAt("resolving");
      logSpy.mockClear();
      sm.reset();
      expect(logSpy).toHaveBeenCalledWith(
        expect.stringContaining("reset → pending"),
      );
      expect(logSpy).toHaveBeenCalledWith(
        expect.stringContaining("was: resolving"),
      );
    });
  });

  // -------------------------------------------------------------------------
  // 9. reset()
  // -------------------------------------------------------------------------

  describe("reset()", () => {
    const allStates: SuggestionState[] = [
      "pending",
      "resolving",
      "accepted",
      "rejected",
      "already-resolved",
      "error",
    ];

    it.each(allStates)("resets to 'pending' from '%s'", (state) => {
      const sm = machineAt(state);
      sm.reset();
      expect(sm.state).toBe("pending");
    });

    it("is idempotent: reset from 'pending' stays 'pending'", () => {
      const sm = new SuggestionStateMachine();
      sm.reset();
      expect(sm.state).toBe("pending");
    });

    it("enables a new accept cycle after reset from error", () => {
      const sm = machineAt("error");
      sm.reset();
      sm.transition("resolving");
      sm.transition("accepted");
      expect(sm.state).toBe("accepted");
    });

    it("enables a new accept cycle after reset from accepted (terminal)", () => {
      const sm = machineAt("accepted");
      sm.reset();
      sm.transition("resolving");
      sm.transition("rejected");
      expect(sm.state).toBe("rejected");
    });
  });

  // -------------------------------------------------------------------------
  // 10. mapResultStatusToState
  // -------------------------------------------------------------------------

  describe("mapResultStatusToState", () => {
    it("maps 'accepted' → 'accepted'", () => {
      expect(mapResultStatusToState("accepted")).toBe("accepted");
    });

    it("maps 'rejected' → 'rejected'", () => {
      expect(mapResultStatusToState("rejected")).toBe("rejected");
    });

    it("maps 'already-resolved' → 'already-resolved'", () => {
      expect(mapResultStatusToState("already-resolved")).toBe(
        "already-resolved",
      );
    });

    it("maps 'cc-not-found' → 'error'", () => {
      expect(mapResultStatusToState("cc-not-found")).toBe("error");
    });

    it("maps 'not-found' → 'error'", () => {
      expect(mapResultStatusToState("not-found")).toBe("error");
    });

    it("maps 'error' → 'error'", () => {
      expect(mapResultStatusToState("error")).toBe("error");
    });
  });
});
