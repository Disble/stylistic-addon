import { describe, expect, it, vi } from "vitest";
import { TrackedChangeResolutionExecutor } from "./TrackedChangeResolutionExecutor";

/**
 * Builds a minimal Word.RequestContext stub whose
 * `document.body.getTrackedChanges()` returns the configured items count on
 * each call. The count function lets a test simulate the document evolving
 * (or not evolving) between executor steps.
 */
function buildBodyTrackedChangeContextStub(
  bodyTrackedChangeCountSequence: Array<number | Error>,
): Word.RequestContext {
  let invocationIndex = 0;
  const sync = vi.fn(async () => undefined);

  const context = {
    sync,
    document: {
      body: {
        getTrackedChanges: vi.fn(() => {
          const probe =
            bodyTrackedChangeCountSequence[
              Math.min(
                invocationIndex,
                bodyTrackedChangeCountSequence.length - 1,
              )
            ] ?? 0;
          invocationIndex += 1;

          if (probe instanceof Error) {
            throw probe;
          }

          return {
            items: Array.from({ length: probe }, (_, index) => ({
              id: `tc-${index}`,
              type: index % 2 === 0 ? "Deleted" : "Added",
            })),
            load: vi.fn(),
          };
        }),
      },
    },
  } as unknown as Word.RequestContext;

  return context;
}

describe("TrackedChangeResolutionExecutor silent no-op detection", () => {
  it("applies accept replace steps in Added then Deleted order", async () => {
    const executor = new TrackedChangeResolutionExecutor("s-accept-order", "accept");
    const context = buildBodyTrackedChangeContextStub([2, 1, 1, 0]);
    const callOrder: string[] = [];

    const report = await executor.apply(context, [
      {
        type: "Deleted",
        accept: vi.fn(() => {
          callOrder.push("accept-deleted");
        }),
      } as unknown as Word.TrackedChange,
      {
        type: "Added",
        accept: vi.fn(() => {
          callOrder.push("accept-added");
        }),
      } as unknown as Word.TrackedChange,
    ]);

    expect(report.error).toBeUndefined();
    expect(report.completed).toBe(2);
    expect(callOrder).toEqual(["accept-added", "accept-deleted"]);
  });

  it("keeps reject replace steps in Deleted then Added order", async () => {
    const executor = new TrackedChangeResolutionExecutor("s-reject-order", "reject");
    const context = buildBodyTrackedChangeContextStub([2, 1, 1, 0]);
    const callOrder: string[] = [];

    const report = await executor.apply(context, [
      {
        type: "Added",
        reject: vi.fn(() => {
          callOrder.push("reject-added");
        }),
      } as unknown as Word.TrackedChange,
      {
        type: "Deleted",
        reject: vi.fn(() => {
          callOrder.push("reject-deleted");
        }),
      } as unknown as Word.TrackedChange,
    ]);

    expect(report.error).toBeUndefined();
    expect(report.completed).toBe(2);
    expect(callOrder).toEqual(["reject-deleted", "reject-added"]);
  });

  it("flags a reject step as silent no-op when bodyTrackedChange count does not decrease after sync", async () => {
    const executor = new TrackedChangeResolutionExecutor("s-1", "reject");
    // Body probes per step:
    //   1. before step 0 (Deleted)
    //   2. after  step 0 (Deleted)  -> NO decrease (3 -> 3): silent no-op
    //   3. before step 1 (Added)
    //   4. after  step 1 (Added)    -> decreases (2 -> 1): real mutation
    const context = buildBodyTrackedChangeContextStub([3, 3, 2, 1]);
    const rejectDeleted = vi.fn();
    const rejectAdded = vi.fn();
    const trackedChanges = [
      { id: "tc-d", type: "Deleted", reject: rejectDeleted } as unknown as Word.TrackedChange,
      { id: "tc-a", type: "Added", reject: rejectAdded } as unknown as Word.TrackedChange,
    ];

    const report = await executor.apply(context, trackedChanges);

    expect(report.attempted).toBe(2);
    expect(report.completed).toBe(2);
    expect(report.error).toBeUndefined();
    expect(report.silentNoOpDetected).toEqual({
      stepIndex: 0,
      trackedChangeType: "Deleted",
      bodyTrackedChangeCountBefore: 3,
      bodyTrackedChangeCountAfter: 3,
    });
    expect(rejectDeleted).toHaveBeenCalledOnce();
    expect(rejectAdded).toHaveBeenCalledOnce();
  });

  it("does not flag silent no-op when bodyTrackedChange count decreases after sync", async () => {
    const executor = new TrackedChangeResolutionExecutor("s-2", "reject");
    const context = buildBodyTrackedChangeContextStub([2, 1, 1, 0]);
    const rejectDeleted = vi.fn();
    const rejectAdded = vi.fn();

    const report = await executor.apply(context, [
      { type: "Deleted", reject: rejectDeleted } as unknown as Word.TrackedChange,
      { type: "Added", reject: rejectAdded } as unknown as Word.TrackedChange,
    ]);

    expect(report.silentNoOpDetected).toBeUndefined();
    expect(report.completed).toBe(2);
  });

  it("flags an accept step as silent no-op when bodyTrackedChange count does not decrease after sync", async () => {
    const executor = new TrackedChangeResolutionExecutor("s-3", "accept");
    const context = buildBodyTrackedChangeContextStub([2, 2, 2, 2]);

    const report = await executor.apply(context, [
      { type: "Deleted", accept: vi.fn() } as unknown as Word.TrackedChange,
      { type: "Added", accept: vi.fn() } as unknown as Word.TrackedChange,
    ]);

    expect(report.silentNoOpDetected).toEqual({
      stepIndex: 0,
      trackedChangeType: "Added",
      bodyTrackedChangeCountBefore: 2,
      bodyTrackedChangeCountAfter: 2,
    });
  });

  it("reports unverified mutation when the body count probe fails instead of treating unknown as zero", async () => {
    const executor = new TrackedChangeResolutionExecutor("s-5", "accept");
    const context = buildBodyTrackedChangeContextStub([
      new Error("InvalidRibbonDefinition"),
      1,
    ]);
    const acceptDeleted = vi.fn();

    const report = await executor.apply(context, [
      { type: "Deleted", accept: acceptDeleted } as unknown as Word.TrackedChange,
    ]);

    expect(report.error).toBeUndefined();
    expect(report.completed).toBe(1);
    expect(report.silentNoOpDetected).toBeUndefined();
    expect(report.unverifiedMutation).toEqual({
      stepIndex: 0,
      trackedChangeType: "Deleted",
      reason: "body-count-probe-failed",
      bodyTrackedChangeCountBeforeError: "InvalidRibbonDefinition",
      bodyTrackedChangeCountAfter: 1,
    });
    expect(acceptDeleted).toHaveBeenCalledOnce();
  });

  it("does not flag silent no-op when bodyTrackedChangeCountBefore is 0 (mocks that do not expose body counts must be tolerated)", async () => {
    const executor = new TrackedChangeResolutionExecutor("s-4", "reject");
    const context = buildBodyTrackedChangeContextStub([0, 0, 0, 0]);

    const report = await executor.apply(context, [
      { type: "Deleted", reject: vi.fn() } as unknown as Word.TrackedChange,
      { type: "Added", reject: vi.fn() } as unknown as Word.TrackedChange,
    ]);

    expect(report.silentNoOpDetected).toBeUndefined();
  });
});
