/* global console */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  ExecuteResolutionStateMachine,
  InvalidExecuteResolutionTransitionError,
} from "./ExecuteResolutionStateMachine";

describe("ExecuteResolutionStateMachine", () => {
  beforeEach(() => {
    vi.spyOn(console, "log").mockImplementation(() => undefined);
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("starts idle with no active phase", () => {
    const machine = new ExecuteResolutionStateMachine();

    expect(machine.state).toBe("idle");
    expect(machine.currentPhase).toBeNull();
    expect(machine.isTerminal).toBe(false);
  });

  it("maps each active state to the expected resolution phase", () => {
    const machine = new ExecuteResolutionStateMachine();

    machine.transition("locating");
    expect(machine.currentPhase).toBe("locate");

    machine.transition("observing-before");
    expect(machine.currentPhase).toBe("observe-before");

    machine.transition("executing");
    expect(machine.currentPhase).toBe("execute");

    machine.transition("cleaning-comment");
    expect(machine.currentPhase).toBe("cleanup-comment");

    machine.transition("cleaning-anchor");
    expect(machine.currentPhase).toBe("cleanup-anchor");

    machine.transition("inspecting-after");
    expect(machine.currentPhase).toBe("inspect-after");
  });

  it("supports the full successful workflow path", () => {
    const machine = new ExecuteResolutionStateMachine();

    machine.transition("locating");
    machine.transition("observing-before");
    machine.transition("executing");
    machine.transition("cleaning-comment");
    machine.transition("cleaning-anchor");
    machine.transition("inspecting-after");
    machine.transition("completed");

    expect(machine.state).toBe("completed");
    expect(machine.isTerminal).toBe(true);
    expect(machine.currentPhase).toBeNull();
  });

  it("supports failure from any active phase", () => {
    const machine = new ExecuteResolutionStateMachine();

    machine.transition("locating");
    machine.fail();

    expect(machine.state).toBe("failed");
    expect(machine.isTerminal).toBe(true);
    expect(machine.currentPhase).toBe("locate");
  });

  it("throws on invalid transition", () => {
    const machine = new ExecuteResolutionStateMachine();

    expect(() => machine.transition("executing")).toThrow(InvalidExecuteResolutionTransitionError);
  });
});
