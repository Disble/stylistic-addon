import type { IAnalysisPort } from "../../domain/ports";
import type { TextChunk } from "../../domain/chunking/TextChunk.types";
import type {
  ChunkPollResult,
  ChunkSubmitResult,
  WorkflowSubmitContext,
} from "../../domain/mastra/MastraWorkflow.types";
import { RetryAnalysisDecorator } from "../RetryAnalysisDecorator";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Creates a mock IAnalysisPort with vi.fn() stubs. */
type MockAnalysisPort = IAnalysisPort & {
  checkConnection: ReturnType<typeof vi.fn<IAnalysisPort["checkConnection"]>>;
  submitChunkAnalysis: ReturnType<typeof vi.fn<IAnalysisPort["submitChunkAnalysis"]>>;
  pollChunkAnalysis: ReturnType<typeof vi.fn<IAnalysisPort["pollChunkAnalysis"]>>;
  cancelChunkAnalysis: ReturnType<typeof vi.fn<IAnalysisPort["cancelChunkAnalysis"]>>;
  retryPollChunkAnalysis: ReturnType<typeof vi.fn<IAnalysisPort["retryPollChunkAnalysis"]>>;
};

/** Creates a mock IAnalysisPort with vi.fn() stubs. */
function createMockPort(): MockAnalysisPort {
  return {
    checkConnection: vi.fn<IAnalysisPort["checkConnection"]>(),
    submitChunkAnalysis: vi.fn<IAnalysisPort["submitChunkAnalysis"]>(),
    pollChunkAnalysis: vi.fn<IAnalysisPort["pollChunkAnalysis"]>(),
    cancelChunkAnalysis: vi.fn<IAnalysisPort["cancelChunkAnalysis"]>(),
    retryPollChunkAnalysis: vi.fn<IAnalysisPort["retryPollChunkAnalysis"]>(),
  };
}

/** A reusable TextChunk fixture. */
function makeChunk(overrides: Partial<TextChunk> = {}): TextChunk {
  return {
    text: "Some text to analyze.",
    index: 0,
    total: 1,
    startOffset: 0,
    ...overrides,
  };
}

function submitSuccess(chunkIndex = 0, runId = `run-${chunkIndex}`): ChunkSubmitResult {
  return { chunkIndex, runId };
}

function submitFailure(chunkIndex = 0, error = "Backend timeout"): ChunkSubmitResult {
  return { chunkIndex, error };
}

function pollRunning(chunkIndex = 0, runId = `run-${chunkIndex}`): ChunkPollResult {
  return { chunkIndex, runId, status: "running", origin: "backend", suggestions: [] };
}

function pollSuccess(chunkIndex = 0, runId = `run-${chunkIndex}`): ChunkPollResult {
  return { chunkIndex, runId, status: "success", origin: "backend", suggestions: [] };
}

function makeSubmitContext(overrides: Partial<WorkflowSubmitContext> = {}): WorkflowSubmitContext {
  return {
    documentUuid: "11111111-1111-4111-8111-111111111111",
    genero: "general",
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("RetryAnalysisDecorator", () => {
  const MAX_RETRIES = 3;
  const BASE_DELAY_MS = 1_000;

  let mockPort: ReturnType<typeof createMockPort>;
  let decorator: RetryAnalysisDecorator;

  let warnSpy: ReturnType<typeof vi.spyOn>;
  let errorSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.useFakeTimers();
    mockPort = createMockPort();
    decorator = new RetryAnalysisDecorator(mockPort, MAX_RETRIES, BASE_DELAY_MS);
    warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    errorSpy = vi.spyOn(console, "error").mockImplementation(() => {});
  });

  afterEach(() => {
    vi.useRealTimers();
    warnSpy.mockRestore();
    errorSpy.mockRestore();
  });

  // -----------------------------------------------------------------------
  // checkConnection — pure passthrough, no retry
  // -----------------------------------------------------------------------
  describe("checkConnection", () => {
    it("delegates to the wrapped port and returns its result", async () => {
      mockPort.checkConnection.mockResolvedValue(true);

      const result = await decorator.checkConnection();

      expect(result).toBe(true);
      expect(mockPort.checkConnection).toHaveBeenCalledOnce();
    });

    it("returns false when the wrapped port returns false", async () => {
      mockPort.checkConnection.mockResolvedValue(false);

      const result = await decorator.checkConnection();

      expect(result).toBe(false);
    });

    it("does NOT retry if checkConnection rejects", async () => {
      mockPort.checkConnection.mockRejectedValue(new Error("Network down"));

      await expect(decorator.checkConnection()).rejects.toThrow("Network down");
      expect(mockPort.checkConnection).toHaveBeenCalledOnce();
    });
  });

  describe("submitChunkAnalysis", () => {
    it("returns the first successful submission without retrying", async () => {
      mockPort.submitChunkAnalysis.mockResolvedValue(submitSuccess(0, "run-123"));

      const result = await decorator.submitChunkAnalysis(makeChunk(), makeSubmitContext());

      expect(result).toEqual({ chunkIndex: 0, runId: "run-123" });
      expect(mockPort.submitChunkAnalysis).toHaveBeenCalledOnce();
      expect(warnSpy).not.toHaveBeenCalled();
    });

    it("retries when the wrapped port returns no runId and later succeeds", async () => {
      mockPort.submitChunkAnalysis
        .mockResolvedValueOnce(submitFailure(0, "temporary submit failure"))
        .mockResolvedValueOnce(submitSuccess(0, "run-456"));

      const promise = decorator.submitChunkAnalysis(makeChunk(), makeSubmitContext());
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);

      await expect(promise).resolves.toEqual({
        chunkIndex: 0,
        runId: "run-456",
      });
      expect(mockPort.submitChunkAnalysis).toHaveBeenCalledTimes(2);
    });

    it("retries thrown submit errors and returns a normalized failure after exhaustion", async () => {
      mockPort.submitChunkAnalysis.mockRejectedValue(new Error("socket hang up"));

      const promise = decorator.submitChunkAnalysis(makeChunk(), makeSubmitContext());
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      await expect(promise).resolves.toEqual({
        chunkIndex: 0,
        error: "Chunk 1: socket hang up",
      });
      expect(mockPort.submitChunkAnalysis).toHaveBeenCalledTimes(MAX_RETRIES + 1);
      expect(errorSpy).toHaveBeenCalled();
    });

    it("does not retry deterministic 4xx submit errors", async () => {
      const error = Object.assign(new Error("HTTP error! status: 400"), {
        status: 400,
      });
      mockPort.submitChunkAnalysis.mockRejectedValue(error);

      await expect(
        decorator.submitChunkAnalysis(makeChunk(), makeSubmitContext())
      ).resolves.toEqual({
        chunkIndex: 0,
        error: "Chunk 1: HTTP error! status: 400",
      });

      expect(mockPort.submitChunkAnalysis).toHaveBeenCalledOnce();
    });
  });

  describe("pollChunkAnalysis", () => {
    it("returns running without retrying because it is not an error", async () => {
      mockPort.pollChunkAnalysis.mockResolvedValue(pollRunning(0, "run-123"));

      const result = await decorator.pollChunkAnalysis(0, "run-123");

      expect(result).toEqual({
        chunkIndex: 0,
        runId: "run-123",
        status: "running",
        origin: "backend",
        suggestions: [],
      });
      expect(mockPort.pollChunkAnalysis).toHaveBeenCalledOnce();
    });

    it("retries thrown poll errors and eventually returns success", async () => {
      mockPort.pollChunkAnalysis
        .mockRejectedValueOnce(new Error("temporary outage"))
        .mockResolvedValueOnce(pollSuccess(0, "run-123"));

      const promise = decorator.pollChunkAnalysis(0, "run-123");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);

      await expect(promise).resolves.toEqual({
        chunkIndex: 0,
        runId: "run-123",
        status: "success",
        origin: "backend",
        suggestions: [],
      });
      expect(mockPort.pollChunkAnalysis).toHaveBeenCalledTimes(2);
    });

    it("returns a failed poll result after exhausting thrown poll errors", async () => {
      mockPort.pollChunkAnalysis.mockRejectedValue(new Error("poll timeout"));

      const promise = decorator.pollChunkAnalysis(4, "run-789");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      await expect(promise).resolves.toEqual({
        chunkIndex: 4,
        runId: "run-789",
        status: "retryable-failure",
        origin: "frontend-retryable",
        suggestions: [],
        error: "Chunk 5: poll timeout",
      });
      expect(mockPort.pollChunkAnalysis).toHaveBeenCalledTimes(MAX_RETRIES + 1);
    });

    it("does not retry deterministic 4xx poll errors", async () => {
      const error = Object.assign(new Error("HTTP error! status: 400"), {
        status: 400,
      });
      mockPort.pollChunkAnalysis.mockRejectedValue(error);

      await expect(decorator.pollChunkAnalysis(4, "run-789")).resolves.toEqual({
        chunkIndex: 4,
        runId: "run-789",
        status: "retryable-failure",
        origin: "frontend-retryable",
        suggestions: [],
        error: "Chunk 5: HTTP error! status: 400",
      });

      expect(mockPort.pollChunkAnalysis).toHaveBeenCalledOnce();
    });

    it("still retries 429 poll errors because they can be transient", async () => {
      const error = Object.assign(new Error("HTTP error! status: 429"), {
        status: 429,
      });
      mockPort.pollChunkAnalysis.mockRejectedValue(error);

      const promise = decorator.pollChunkAnalysis(4, "run-789");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      await expect(promise).resolves.toEqual({
        chunkIndex: 4,
        runId: "run-789",
        status: "retryable-failure",
        origin: "frontend-retryable",
        suggestions: [],
        error: "Chunk 5: HTTP error! status: 429",
      });

      expect(mockPort.pollChunkAnalysis).toHaveBeenCalledTimes(MAX_RETRIES + 1);
    });

    it("delegates cancelChunkAnalysis without adding retry behavior", async () => {
      mockPort.cancelChunkAnalysis.mockResolvedValue({
        chunkIndex: 2,
        runId: "run-2",
        canceled: true,
      });

      await expect(decorator.cancelChunkAnalysis(2, "run-2")).resolves.toEqual({
        chunkIndex: 2,
        runId: "run-2",
        canceled: true,
      });

      expect(mockPort.cancelChunkAnalysis).toHaveBeenCalledWith(2, "run-2");
    });
  });
});
