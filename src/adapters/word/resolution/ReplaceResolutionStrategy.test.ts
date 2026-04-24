import { describe, expect, it } from "vitest";
import {
  AcceptReplaceResolutionStrategy,
  RejectReplaceResolutionStrategy,
} from "./ReplaceResolutionStrategyContext";

describe("ReplaceResolutionStrategy", () => {
  it("delegates accept replace policy to the concrete accept strategy", () => {
    const strategy = new AcceptReplaceResolutionStrategy();

    expect(strategy.actionLabel).toBe("aceptación");
    expect(strategy.semanticOrder).toEqual(["Added", "Deleted"]);
    expect(strategy.priorityFor("Added")).toBe(0);
    expect(strategy.priorityFor("Deleted")).toBe(1);
    expect(strategy.priorityFor("Formatting")).toBe(2);
  });

  it("delegates reject replace policy to the concrete reject strategy", () => {
    const strategy = new RejectReplaceResolutionStrategy();

    expect(strategy.actionLabel).toBe("rechazo");
    expect(strategy.semanticOrder).toEqual(["Deleted", "Added"]);
    expect(strategy.priorityFor("Deleted")).toBe(0);
    expect(strategy.priorityFor("Added")).toBe(1);
    expect(strategy.priorityFor("Formatting")).toBe(2);
  });
});
