import { describe, it, expect, vi, beforeEach } from "vitest";
import { MockFeedbackAdapter } from "./MockFeedbackAdapter";
import type { FeedbackPayload } from "../../domain/types";

function makePayload(overrides: Partial<FeedbackPayload> = {}): FeedbackPayload {
  return {
    category: "Muletilla",
    originalText: "básicamente",
    suggestedText: "",
    justification: "Filler word.",
    rating: "negative",
    severity: "medium",
    ...overrides,
  };
}

describe("MockFeedbackAdapter", () => {
  let logSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    logSpy = vi.spyOn(console, "log").mockImplementation(() => {});
  });

  afterEach(() => {
    logSpy.mockRestore();
  });

  it("returns Promise<void> (resolves without throwing)", async () => {
    const adapter = new MockFeedbackAdapter();
    await expect(adapter.sendFeedback(makePayload())).resolves.toBeUndefined();
  });

  it("logs the payload via console.log", async () => {
    const adapter = new MockFeedbackAdapter();
    const payload = makePayload({ comment: "Muy bien" });
    await adapter.sendFeedback(payload);

    expect(logSpy).toHaveBeenCalledWith("[MockFeedbackAdapter] sendFeedback:", payload);
  });

  it("logs payload including justification", async () => {
    const adapter = new MockFeedbackAdapter();
    const payload = makePayload({ justification: "Palabra de relleno" });
    await adapter.sendFeedback(payload);

    expect(logSpy).toHaveBeenCalledWith(
      "[MockFeedbackAdapter] sendFeedback:",
      expect.objectContaining({ justification: "Palabra de relleno" })
    );
  });
});
