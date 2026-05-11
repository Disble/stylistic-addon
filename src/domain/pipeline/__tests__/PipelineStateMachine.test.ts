import { describe, expect, it, vi } from "vitest";
import { PipelineStateMachine } from "../PipelineStateMachine";

describe("PipelineStateMachine", () => {
  it("starts idle and reports not running", () => {
    const machine = new PipelineStateMachine();

    expect(machine.state).toBe("idle");
    expect(machine.isRunning).toBe(false);
    expect(machine.canTransition("reading")).toBe(true);
  });

  it("supports the happy path from idle to done", () => {
    const machine = new PipelineStateMachine();

    machine.transition("reading");
    machine.transition("connecting");
    machine.transition("chunking");
    machine.transition("analyzing");
    machine.transition("applying");
    machine.transition("done");

    expect(machine.state).toBe("done");
    expect(machine.isRunning).toBe(false);
  });

  it("allows cancellation and error exits from in-flight states", () => {
    const cancelled = new PipelineStateMachine();
    cancelled.transition("reading");
    cancelled.transition("idle");

    const failed = new PipelineStateMachine();
    failed.transition("reading");
    failed.transition("connecting");
    failed.transition("error");

    expect(cancelled.state).toBe("idle");
    expect(failed.state).toBe("error");
    expect(failed.isRunning).toBe(false);
  });

  it("throws and preserves state on invalid transitions", () => {
    const machine = new PipelineStateMachine();

    expect(() => machine.transition("done")).toThrow(
      '[PipelineStateMachine] Invalid transition: "idle" → "done". Allowed: [reading]'
    );
    expect(machine.state).toBe("idle");
    expect(machine.canTransition("done")).toBe(false);
  });

  it("resets to idle from terminal and mid-pipeline states", () => {
    const fromDone = new PipelineStateMachine();
    fromDone.transition("reading");
    fromDone.transition("connecting");
    fromDone.transition("chunking");
    fromDone.transition("analyzing");
    fromDone.transition("applying");
    fromDone.transition("done");
    fromDone.reset();

    const fromAnalyzing = new PipelineStateMachine();
    fromAnalyzing.transition("reading");
    fromAnalyzing.transition("connecting");
    fromAnalyzing.transition("chunking");
    fromAnalyzing.transition("analyzing");
    fromAnalyzing.reset();

    expect(fromDone.state).toBe("idle");
    expect(fromAnalyzing.state).toBe("idle");
  });

  it("logs transitions and resets", () => {
    const logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
    const machine = new PipelineStateMachine();

    machine.transition("reading");
    machine.reset();

    expect(logSpy).toHaveBeenNthCalledWith(1, expect.stringContaining("idle → reading"));
    expect(logSpy).toHaveBeenNthCalledWith(
      2,
      expect.stringContaining("reset → idle (was: reading)")
    );

    logSpy.mockRestore();
  });
});
