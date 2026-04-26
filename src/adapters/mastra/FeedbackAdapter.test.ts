/* global console */

import { beforeEach, describe, expect, it, vi } from "vitest";
import type { FeedbackPayload } from "../../domain/suggestion/SuggestionResolutionWorkflow.types";

// ---------------------------------------------------------------------------
// Mock @mastra/client-js using vi.hoisted so mocks are available at module init time
// ---------------------------------------------------------------------------

const feedbackAdapterMocks = vi.hoisted(() => ({
  start: vi.fn(),
  createRun: vi.fn(),
  getWorkflow: vi.fn(),
  clientConstructor: vi.fn(),
}));

vi.mock("@mastra/client-js", () => ({
  MastraClient: class {
    constructor(...args: unknown[]) {
      feedbackAdapterMocks.clientConstructor(...args);
    }
    getWorkflow(id: string) {
      return feedbackAdapterMocks.getWorkflow(id);
    }
  },
}));

import { FEEDBACK_WORKFLOW_ID } from "../../infrastructure/config";
import { FeedbackAdapter } from "./FeedbackAdapter";

function makePayload(
  overrides: Partial<FeedbackPayload> = {},
): FeedbackPayload {
  return {
    autorSlug: "disble",
    category: "Redundancia",
    context: "Frase con completamente necesario.",
    anchor: "completamente necesario",
    suggestedText: "necesario",
    justification: "Ya implica completitud.",
    action: "accept",
    severity: "high",
    suggestionType: "track-change",
    ...overrides,
  };
}

describe("FeedbackAdapter", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    feedbackAdapterMocks.start.mockResolvedValue(undefined);
    feedbackAdapterMocks.createRun.mockResolvedValue({
      runId: "run-123",
      start: feedbackAdapterMocks.start,
    });
    feedbackAdapterMocks.getWorkflow.mockReturnValue({
      createRun: feedbackAdapterMocks.createRun,
    });
  });

  it("calls getWorkflow with FEEDBACK_WORKFLOW_ID", async () => {
    const adapter = new FeedbackAdapter();
    await adapter.sendFeedback(makePayload());

    expect(feedbackAdapterMocks.getWorkflow).toHaveBeenCalledWith(
      FEEDBACK_WORKFLOW_ID,
    );
  });

  it("calls createRun() then run.start() with the payload as inputData", async () => {
    const adapter = new FeedbackAdapter();
    const payload = makePayload({
      action: "reject",
      comment: "test comment",
    });
    await adapter.sendFeedback(payload);

    expect(feedbackAdapterMocks.createRun).toHaveBeenCalledOnce();
    expect(feedbackAdapterMocks.start).toHaveBeenCalledWith({
      inputData: payload,
    });
  });

  it("includes justification in the payload sent to start()", async () => {
    const adapter = new FeedbackAdapter();
    const payload = makePayload({ justification: "Frase innecesaria" });
    await adapter.sendFeedback(payload);

    expect(feedbackAdapterMocks.start).toHaveBeenCalledWith({
      inputData: expect.objectContaining({
        justification: "Frase innecesaria",
      }),
    });
  });

  it("swallows errors silently — never throws when createRun rejects", async () => {
    feedbackAdapterMocks.createRun.mockRejectedValue(
      new Error("Network error"),
    );

    const adapter = new FeedbackAdapter();
    // Must not throw
    await expect(adapter.sendFeedback(makePayload())).resolves.toBeUndefined();
  });

  it("swallows errors silently — never throws when run.start() rejects", async () => {
    feedbackAdapterMocks.start.mockRejectedValue(new Error("Backend 500"));

    const adapter = new FeedbackAdapter();
    await expect(adapter.sendFeedback(makePayload())).resolves.toBeUndefined();
  });
});
