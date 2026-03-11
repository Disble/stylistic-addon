import { PipelineStateMachine } from "./PipelineStateMachine";
import type { PipelineState } from "../types";

// All states for exhaustive iteration
const ALL_STATES: PipelineState[] = [
  "idle",
  "reading",
  "connecting",
  "chunking",
  "analyzing",
  "applying",
  "done",
  "error",
];

/**
 * Helper: drives a fresh state machine to the desired state via valid
 * transitions so we can test behavior FROM that state.
 */
function machineAt(target: PipelineState): PipelineStateMachine {
  const sm = new PipelineStateMachine();
  const path: Record<PipelineState, PipelineState[]> = {
    idle: [],
    reading: ["reading"],
    connecting: ["reading", "connecting"],
    chunking: ["reading", "connecting", "chunking"],
    analyzing: ["reading", "connecting", "chunking", "analyzing"],
    applying: ["reading", "connecting", "chunking", "analyzing", "applying"],
    done: ["reading", "connecting", "chunking", "analyzing", "applying", "done"],
    error: ["reading", "error"],
  };
  for (const step of path[target]) {
    sm.transition(step);
  }
  return sm;
}

describe("PipelineStateMachine", () => {
  let consoleSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    consoleSpy = vi.spyOn(console, "log").mockImplementation(() => {});
  });

  afterEach(() => {
    consoleSpy.mockRestore();
  });

  // -----------------------------------------------------------------------
  // Initial state
  // -----------------------------------------------------------------------
  describe("initial state", () => {
    it("starts in idle", () => {
      const sm = new PipelineStateMachine();
      expect(sm.state).toBe("idle");
    });

    it("is not running when idle", () => {
      const sm = new PipelineStateMachine();
      expect(sm.isRunning).toBe(false);
    });
  });

  // -----------------------------------------------------------------------
  // Valid transitions (happy path)
  // -----------------------------------------------------------------------
  describe("valid transitions", () => {
    it("idle → reading", () => {
      const sm = machineAt("idle");
      sm.transition("reading");
      expect(sm.state).toBe("reading");
    });

    it("reading → connecting", () => {
      const sm = machineAt("reading");
      sm.transition("connecting");
      expect(sm.state).toBe("connecting");
    });

    it("reading → idle (cancel during read)", () => {
      const sm = machineAt("reading");
      sm.transition("idle");
      expect(sm.state).toBe("idle");
    });

    it("reading → error", () => {
      const sm = machineAt("reading");
      sm.transition("error");
      expect(sm.state).toBe("error");
    });

    it("connecting → chunking", () => {
      const sm = machineAt("connecting");
      sm.transition("chunking");
      expect(sm.state).toBe("chunking");
    });

    it("connecting → idle (cancel during connect)", () => {
      const sm = machineAt("connecting");
      sm.transition("idle");
      expect(sm.state).toBe("idle");
    });

    it("connecting → error", () => {
      const sm = machineAt("connecting");
      sm.transition("error");
      expect(sm.state).toBe("error");
    });

    it("chunking → analyzing", () => {
      const sm = machineAt("chunking");
      sm.transition("analyzing");
      expect(sm.state).toBe("analyzing");
    });

    it("chunking → idle (cancel during chunking)", () => {
      const sm = machineAt("chunking");
      sm.transition("idle");
      expect(sm.state).toBe("idle");
    });

    it("chunking → error", () => {
      const sm = machineAt("chunking");
      sm.transition("error");
      expect(sm.state).toBe("error");
    });

    it("analyzing → applying", () => {
      const sm = machineAt("analyzing");
      sm.transition("applying");
      expect(sm.state).toBe("applying");
    });

    it("analyzing → idle (cancel during analysis)", () => {
      const sm = machineAt("analyzing");
      sm.transition("idle");
      expect(sm.state).toBe("idle");
    });

    it("analyzing → error", () => {
      const sm = machineAt("analyzing");
      sm.transition("error");
      expect(sm.state).toBe("error");
    });

    it("applying → done", () => {
      const sm = machineAt("applying");
      sm.transition("done");
      expect(sm.state).toBe("done");
    });

    it("applying → idle (cancel during apply)", () => {
      const sm = machineAt("applying");
      sm.transition("idle");
      expect(sm.state).toBe("idle");
    });

    it("applying → error", () => {
      const sm = machineAt("applying");
      sm.transition("error");
      expect(sm.state).toBe("error");
    });

    it("done → idle", () => {
      const sm = machineAt("done");
      sm.transition("idle");
      expect(sm.state).toBe("idle");
    });

    it("error → idle", () => {
      const sm = machineAt("error");
      sm.transition("idle");
      expect(sm.state).toBe("idle");
    });

    it("completes a full happy-path pipeline run", () => {
      const sm = new PipelineStateMachine();
      sm.transition("reading");
      sm.transition("connecting");
      sm.transition("chunking");
      sm.transition("analyzing");
      sm.transition("applying");
      sm.transition("done");
      expect(sm.state).toBe("done");
    });
  });

  // -----------------------------------------------------------------------
  // Invalid transitions
  // -----------------------------------------------------------------------
  describe("invalid transitions", () => {
    it("throws on idle → idle (self-transition)", () => {
      const sm = machineAt("idle");
      expect(() => sm.transition("idle")).toThrow(/Invalid transition/);
    });

    it("throws on idle → connecting (skipping reading)", () => {
      const sm = machineAt("idle");
      expect(() => sm.transition("connecting")).toThrow(/Invalid transition/);
    });

    it("throws on idle → chunking", () => {
      const sm = machineAt("idle");
      expect(() => sm.transition("chunking")).toThrow(/Invalid transition/);
    });

    it("throws on idle → analyzing", () => {
      const sm = machineAt("idle");
      expect(() => sm.transition("analyzing")).toThrow(/Invalid transition/);
    });

    it("throws on idle → applying", () => {
      const sm = machineAt("idle");
      expect(() => sm.transition("applying")).toThrow(/Invalid transition/);
    });

    it("throws on idle → done", () => {
      const sm = machineAt("idle");
      expect(() => sm.transition("done")).toThrow(/Invalid transition/);
    });

    it("throws on idle → error", () => {
      const sm = machineAt("idle");
      expect(() => sm.transition("error")).toThrow(/Invalid transition/);
    });

    it("throws on reading → reading (self-transition)", () => {
      const sm = machineAt("reading");
      expect(() => sm.transition("reading")).toThrow(/Invalid transition/);
    });

    it("throws on reading → chunking (skipping connecting)", () => {
      const sm = machineAt("reading");
      expect(() => sm.transition("chunking")).toThrow(/Invalid transition/);
    });

    it("throws on reading → done", () => {
      const sm = machineAt("reading");
      expect(() => sm.transition("done")).toThrow(/Invalid transition/);
    });

    it("throws on done → done (self-transition)", () => {
      const sm = machineAt("done");
      expect(() => sm.transition("done")).toThrow(/Invalid transition/);
    });

    it("throws on done → reading (must go through idle)", () => {
      const sm = machineAt("done");
      expect(() => sm.transition("reading")).toThrow(/Invalid transition/);
    });

    it("throws on error → error (self-transition)", () => {
      const sm = machineAt("error");
      expect(() => sm.transition("error")).toThrow(/Invalid transition/);
    });

    it("throws on error → reading (must go through idle)", () => {
      const sm = machineAt("error");
      expect(() => sm.transition("reading")).toThrow(/Invalid transition/);
    });

    it("includes current state and target in error message", () => {
      const sm = machineAt("idle");
      expect(() => sm.transition("done")).toThrow('"idle" → "done"');
    });

    it("includes allowed transitions in error message", () => {
      const sm = machineAt("idle");
      expect(() => sm.transition("done")).toThrow("Allowed: [reading]");
    });

    it("does not change state when transition throws", () => {
      const sm = machineAt("idle");
      try {
        sm.transition("done");
      } catch {
        // expected
      }
      expect(sm.state).toBe("idle");
    });
  });

  // -----------------------------------------------------------------------
  // Exhaustive invalid transition matrix
  // -----------------------------------------------------------------------
  describe("exhaustive invalid transitions", () => {
    const validTransitions: Record<PipelineState, PipelineState[]> = {
      idle: ["reading"],
      reading: ["connecting", "idle", "error"],
      connecting: ["chunking", "idle", "error"],
      chunking: ["analyzing", "idle", "error"],
      analyzing: ["applying", "idle", "error"],
      applying: ["done", "idle", "error"],
      done: ["idle"],
      error: ["idle"],
    };

    for (const from of ALL_STATES) {
      for (const to of ALL_STATES) {
        if (!validTransitions[from].includes(to)) {
          it(`throws on ${from} → ${to}`, () => {
            const sm = machineAt(from);
            expect(() => sm.transition(to)).toThrow(/Invalid transition/);
          });
        }
      }
    }
  });

  // -----------------------------------------------------------------------
  // canTransition()
  // -----------------------------------------------------------------------
  describe("canTransition", () => {
    it("returns true for valid transitions", () => {
      const sm = new PipelineStateMachine();
      expect(sm.canTransition("reading")).toBe(true);
    });

    it("returns false for invalid transitions", () => {
      const sm = new PipelineStateMachine();
      expect(sm.canTransition("done")).toBe(false);
    });

    it("does not change state when called", () => {
      const sm = new PipelineStateMachine();
      sm.canTransition("reading");
      expect(sm.state).toBe("idle");
    });

    it("reflects the current state accurately after transitions", () => {
      const sm = new PipelineStateMachine();
      sm.transition("reading");
      expect(sm.canTransition("connecting")).toBe(true);
      expect(sm.canTransition("reading")).toBe(false);
    });

    it("returns true for all valid targets from each state", () => {
      const validTransitions: Record<PipelineState, PipelineState[]> = {
        idle: ["reading"],
        reading: ["connecting", "idle", "error"],
        connecting: ["chunking", "idle", "error"],
        chunking: ["analyzing", "idle", "error"],
        analyzing: ["applying", "idle", "error"],
        applying: ["done", "idle", "error"],
        done: ["idle"],
        error: ["idle"],
      };

      for (const from of ALL_STATES) {
        const sm = machineAt(from);
        for (const to of validTransitions[from]) {
          expect(sm.canTransition(to)).toBe(true);
        }
      }
    });
  });

  // -----------------------------------------------------------------------
  // isRunning
  // -----------------------------------------------------------------------
  describe("isRunning", () => {
    it("is false when idle", () => {
      expect(machineAt("idle").isRunning).toBe(false);
    });

    it("is true when reading", () => {
      expect(machineAt("reading").isRunning).toBe(true);
    });

    it("is true when connecting", () => {
      expect(machineAt("connecting").isRunning).toBe(true);
    });

    it("is true when chunking", () => {
      expect(machineAt("chunking").isRunning).toBe(true);
    });

    it("is true when analyzing", () => {
      expect(machineAt("analyzing").isRunning).toBe(true);
    });

    it("is true when applying", () => {
      expect(machineAt("applying").isRunning).toBe(true);
    });

    it("is false when done", () => {
      expect(machineAt("done").isRunning).toBe(false);
    });

    it("is false when error", () => {
      expect(machineAt("error").isRunning).toBe(false);
    });
  });

  // -----------------------------------------------------------------------
  // reset()
  // -----------------------------------------------------------------------
  describe("reset", () => {
    it("resets to idle from every state", () => {
      for (const state of ALL_STATES) {
        const sm = machineAt(state);
        sm.reset();
        expect(sm.state).toBe("idle");
      }
    });

    it("is idempotent (resetting idle stays idle)", () => {
      const sm = new PipelineStateMachine();
      sm.reset();
      expect(sm.state).toBe("idle");
    });

    it("allows a new pipeline run after reset from done", () => {
      const sm = machineAt("done");
      sm.reset();
      sm.transition("reading");
      expect(sm.state).toBe("reading");
    });

    it("allows a new pipeline run after reset from error", () => {
      const sm = machineAt("error");
      sm.reset();
      sm.transition("reading");
      expect(sm.state).toBe("reading");
    });

    it("matches the transition model by allowing idle from every non-idle state", () => {
      for (const state of ALL_STATES.filter((candidate) => candidate !== "idle")) {
        const sm = machineAt(state);
        expect(sm.canTransition("idle")).toBe(true);
      }
    });
  });

  // -----------------------------------------------------------------------
  // Console logging
  // -----------------------------------------------------------------------
  describe("console logging", () => {
    it("logs on valid transition", () => {
      const sm = new PipelineStateMachine();
      sm.transition("reading");
      expect(consoleSpy).toHaveBeenCalledWith(
        expect.stringContaining("idle → reading")
      );
    });

    it("logs the transition emoji marker", () => {
      const sm = new PipelineStateMachine();
      sm.transition("reading");
      expect(consoleSpy).toHaveBeenCalledWith(
        expect.stringContaining("[StateMachine]")
      );
    });

    it("does not log on invalid transition (throws before logging)", () => {
      const sm = new PipelineStateMachine();
      consoleSpy.mockClear();
      try {
        sm.transition("done");
      } catch {
        // expected
      }
      expect(consoleSpy).not.toHaveBeenCalled();
    });

    it("logs on reset with the previous state", () => {
      const sm = machineAt("analyzing");
      consoleSpy.mockClear();
      sm.reset();
      expect(consoleSpy).toHaveBeenCalledWith(
        expect.stringContaining("reset → idle (was: analyzing)")
      );
    });

    it("logs on reset from idle", () => {
      const sm = new PipelineStateMachine();
      consoleSpy.mockClear();
      sm.reset();
      expect(consoleSpy).toHaveBeenCalledWith(
        expect.stringContaining("reset → idle (was: idle)")
      );
    });
  });

  // -----------------------------------------------------------------------
  // Edge cases
  // -----------------------------------------------------------------------
  describe("edge cases", () => {
    it("supports multiple full pipeline cycles", () => {
      const sm = new PipelineStateMachine();

      // First cycle
      sm.transition("reading");
      sm.transition("connecting");
      sm.transition("chunking");
      sm.transition("analyzing");
      sm.transition("applying");
      sm.transition("done");
      sm.transition("idle");

      // Second cycle
      sm.transition("reading");
      sm.transition("connecting");
      sm.transition("chunking");
      sm.transition("analyzing");
      sm.transition("applying");
      sm.transition("done");

      expect(sm.state).toBe("done");
    });

    it("supports error-recovery-retry cycle", () => {
      const sm = new PipelineStateMachine();

      // First attempt: fails at analyzing
      sm.transition("reading");
      sm.transition("connecting");
      sm.transition("chunking");
      sm.transition("analyzing");
      sm.transition("error");
      sm.transition("idle");

      // Retry: succeeds
      sm.transition("reading");
      sm.transition("connecting");
      sm.transition("chunking");
      sm.transition("analyzing");
      sm.transition("applying");
      sm.transition("done");

      expect(sm.state).toBe("done");
    });

    it("supports cancel-and-restart cycle", () => {
      const sm = new PipelineStateMachine();

      // Start and cancel during reading
      sm.transition("reading");
      sm.transition("idle");

      // Restart full pipeline
      sm.transition("reading");
      sm.transition("connecting");
      sm.transition("chunking");
      sm.transition("analyzing");
      sm.transition("applying");
      sm.transition("done");

      expect(sm.state).toBe("done");
    });

    it("supports reset-and-restart from mid-pipeline", () => {
      const sm = new PipelineStateMachine();

      sm.transition("reading");
      sm.transition("connecting");
      sm.transition("chunking");
      sm.reset(); // force-reset from chunking

      sm.transition("reading");
      sm.transition("connecting");
      sm.transition("chunking");
      sm.transition("analyzing");
      sm.transition("applying");
      sm.transition("done");

      expect(sm.state).toBe("done");
    });
  });
});
