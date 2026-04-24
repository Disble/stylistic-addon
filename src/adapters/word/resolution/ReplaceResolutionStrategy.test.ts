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
  });

  it("delegates reject replace policy to the concrete reject strategy", () => {
    const strategy = new RejectReplaceResolutionStrategy();

    expect(strategy.actionLabel).toBe("rechazo");
    expect(strategy.semanticOrder).toEqual(["Deleted", "Added"]);
  });
});
