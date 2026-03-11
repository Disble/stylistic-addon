import { RetryAnalysisDecorator } from "./RetryAnalysisDecorator";
import type { IAnalysisPort } from "../domain/ports";
import type { TextChunk, ChunkResult, Suggestion } from "../domain/types";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Creates a mock IAnalysisPort with vi.fn() stubs. */
function createMockPort(): { [K in keyof IAnalysisPort]: ReturnType<typeof vi.fn> } {
  return {
    checkConnection: vi.fn(),
    analyzeChunk: vi.fn(),
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

/** A successful ChunkResult (no error field). */
function successResult(chunkIndex = 0, suggestions: Suggestion[] = []): ChunkResult {
  return { chunkIndex, suggestions };
}

/** A failed ChunkResult with an error message. */
function errorResult(chunkIndex = 0, error = "Backend timeout"): ChunkResult {
  return { chunkIndex, suggestions: [], error };
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
    decorator = new RetryAnalysisDecorator(
      mockPort as unknown as IAnalysisPort,
      MAX_RETRIES,
      BASE_DELAY_MS,
    );
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

  // -----------------------------------------------------------------------
  // analyzeChunk — happy path (first attempt succeeds)
  // -----------------------------------------------------------------------
  describe("analyzeChunk — happy path", () => {
    it("returns the result from the wrapped port on first success", async () => {
      const expected = successResult(0, [
        {
          id: "chunk0-0",
          originalText: "foo",
          suggestedText: "bar",
          justification: "clarity",
          category: "Estilo",
          severity: "low",
        },
      ]);
      mockPort.analyzeChunk.mockResolvedValue(expected);

      const result = await decorator.analyzeChunk(makeChunk(), "general", "es");

      expect(result).toEqual(expected);
    });

    it("calls the wrapped port exactly once on success", async () => {
      mockPort.analyzeChunk.mockResolvedValue(successResult());

      await decorator.analyzeChunk(makeChunk(), "general", "es");

      expect(mockPort.analyzeChunk).toHaveBeenCalledOnce();
    });

    it("passes chunk, profile, and language to the wrapped port", async () => {
      const chunk = makeChunk({ index: 5, total: 10 });
      mockPort.analyzeChunk.mockResolvedValue(successResult(5));

      await decorator.analyzeChunk(chunk, "formal", "en");

      expect(mockPort.analyzeChunk).toHaveBeenCalledWith(chunk, "formal", "en");
    });

    it("returns the result unmodified (referential identity)", async () => {
      const expected = successResult(0);
      mockPort.analyzeChunk.mockResolvedValue(expected);

      const result = await decorator.analyzeChunk(makeChunk(), "general", "es");

      expect(result).toBe(expected);
    });

    it("does not log any warnings or errors on first success", async () => {
      mockPort.analyzeChunk.mockResolvedValue(successResult());

      await decorator.analyzeChunk(makeChunk(), "general", "es");

      expect(warnSpy).not.toHaveBeenCalled();
      expect(errorSpy).not.toHaveBeenCalled();
    });
  });

  // -----------------------------------------------------------------------
  // analyzeChunk — retry then succeed
  // -----------------------------------------------------------------------
  describe("analyzeChunk — retry then succeed", () => {
    it("retries and returns success after 1 failure", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "Timeout"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      // After first call fails, decorator delays baseDelayMs * 2^0 = 1000ms
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);

      const result = await promise;

      expect(result).toEqual(successResult(0));
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(2);
    });

    it("retries and returns success after 2 failures", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "Timeout"))
        .mockResolvedValueOnce(errorResult(0, "503 Service Unavailable"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      // Retry 1: delay = 1000ms (base * 2^0)
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      // Retry 2: delay = 2000ms (base * 2^1)
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);

      const result = await promise;

      expect(result).toEqual(successResult(0));
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(3);
    });

    it("retries and returns success after maxRetries failures", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "err1"))
        .mockResolvedValueOnce(errorResult(0, "err2"))
        .mockResolvedValueOnce(errorResult(0, "err3"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      // Retry 1: 1000ms, Retry 2: 2000ms, Retry 3: 4000ms
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      const result = await promise;

      expect(result).toEqual(successResult(0));
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(4); // 1 initial + 3 retries
    });
  });

  // -----------------------------------------------------------------------
  // analyzeChunk — exhaustion (all retries fail)
  // -----------------------------------------------------------------------
  describe("analyzeChunk — exhaustion", () => {
    it("returns an error result after exhausting all retries", async () => {
      mockPort.analyzeChunk.mockResolvedValue(errorResult(0, "Server Error"));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      // Advance through all backoff delays: 1000 + 2000 + 4000 = 7000ms
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      const result = await promise;

      expect(result).toEqual({
        chunkIndex: 0,
        suggestions: [],
        error: "Chunk 1: Server Error",
      });
    });

    it("calls the wrapped port maxRetries + 1 times total", async () => {
      mockPort.analyzeChunk.mockResolvedValue(errorResult(0, "fail"));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      await promise;

      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(MAX_RETRIES + 1);
    });

    it("never throws — always returns a ChunkResult", async () => {
      mockPort.analyzeChunk.mockResolvedValue(errorResult(2, "Boom"));

      const chunk = makeChunk({ index: 2 });
      const promise = decorator.analyzeChunk(chunk, "general", "es");

      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      const result = await promise;

      expect(result.error).toBeDefined();
      expect(result.suggestions).toEqual([]);
    });

    it("formats the error message as 'Chunk {index+1}: {lastError}'", async () => {
      const chunk = makeChunk({ index: 4 });
      mockPort.analyzeChunk.mockResolvedValue(errorResult(4, "rate limited"));

      const promise = decorator.analyzeChunk(chunk, "general", "es");

      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      const result = await promise;

      expect(result.error).toBe("Chunk 5: rate limited");
    });

    it("uses the LAST error message, not the first", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "first error"))
        .mockResolvedValueOnce(errorResult(0, "second error"))
        .mockResolvedValueOnce(errorResult(0, "third error"))
        .mockResolvedValueOnce(errorResult(0, "final error"));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      const result = await promise;

      expect(result.error).toBe("Chunk 1: final error");
    });

    it("retries thrown exceptions and returns a normalized error after exhaustion", async () => {
      mockPort.analyzeChunk.mockRejectedValue(new Error("socket hang up"));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      const result = await promise;

      expect(result).toEqual({
        chunkIndex: 0,
        suggestions: [],
        error: "Chunk 1: socket hang up",
      });
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(MAX_RETRIES + 1);
    });
  });

  // -----------------------------------------------------------------------
  // Exponential backoff timing
  // -----------------------------------------------------------------------
  describe("exponential backoff timing", () => {
    it("applies delay = baseDelayMs * 2^0 on first retry", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "fail"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      // Not enough time — second call should NOT have happened yet
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS - 1);
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(1);

      // Advance to exactly baseDelayMs
      await vi.advanceTimersByTimeAsync(1);
      await promise;
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(2);
    });

    it("applies delay = baseDelayMs * 2^1 on second retry", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "fail"))
        .mockResolvedValueOnce(errorResult(0, "fail"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      // First retry: 1000ms
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(2);

      // Second retry: 2000ms — not there yet at 1999ms
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2 - 1);
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(2);

      // Now advance the last ms
      await vi.advanceTimersByTimeAsync(1);
      await promise;
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(3);
    });

    it("applies delay = baseDelayMs * 2^2 on third retry", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "fail"))
        .mockResolvedValueOnce(errorResult(0, "fail"))
        .mockResolvedValueOnce(errorResult(0, "fail"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      // Retry 1: 1000ms
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(2);

      // Retry 2: 2000ms
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(3);

      // Retry 3: 4000ms — not there yet
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4 - 1);
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(3);

      await vi.advanceTimersByTimeAsync(1);
      await promise;
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(4);
    });

    it("does NOT delay on the first attempt (attempt 0)", async () => {
      mockPort.analyzeChunk.mockResolvedValue(successResult());

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      // The first call should happen synchronously (no timer needed)
      // Just drain microtasks
      await vi.advanceTimersByTimeAsync(0);
      await promise;

      expect(mockPort.analyzeChunk).toHaveBeenCalledOnce();
    });
  });

  // -----------------------------------------------------------------------
  // Console logging behavior
  // -----------------------------------------------------------------------
  describe("console logging", () => {
    it("logs a console.warn with retry info before each retry", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "fail"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await promise;

      // Two warn calls: one for the failure log, one for the retry-about-to-happen log
      // Line 58: retry announce, Line 70-71: failure log
      expect(warnSpy).toHaveBeenCalledWith(
        expect.stringContaining("[RetryDecorator]"),
      );
      expect(warnSpy).toHaveBeenCalledWith(
        expect.stringContaining("Retry 1/3"),
      );
    });

    it("logs a console.warn for each failed attempt", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "err1"))
        .mockResolvedValueOnce(errorResult(0, "err2"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await promise;

      // Each failed attempt logs a warn with "falló"
      const failureLogs = warnSpy.mock.calls.filter(
        (args) => typeof args[0] === "string" && args[0].includes("falló"),
      );
      expect(failureLogs).toHaveLength(2);
    });

    it("logs console.error when all retries are exhausted", async () => {
      mockPort.analyzeChunk.mockResolvedValue(errorResult(0, "persist"));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");

      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      await promise;

      expect(errorSpy).toHaveBeenCalledWith(
        expect.stringContaining("agotó"),
      );
      expect(errorSpy).toHaveBeenCalledWith(
        expect.stringContaining("persist"),
      );
    });

    it("does NOT log console.error when a retry eventually succeeds", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "fail"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await promise;

      expect(errorSpy).not.toHaveBeenCalled();
    });

    it("includes the chunk index in retry log messages", async () => {
      const chunk = makeChunk({ index: 7 });
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(7, "fail"))
        .mockResolvedValueOnce(successResult(7));

      const promise = decorator.analyzeChunk(chunk, "general", "es");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await promise;

      expect(warnSpy).toHaveBeenCalledWith(
        expect.stringContaining("chunk #7"),
      );
    });

    it("includes the delay in the retry log message", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "fail"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await promise;

      // First retry delay = 1000ms
      expect(warnSpy).toHaveBeenCalledWith(
        expect.stringContaining("1000ms"),
      );
    });
  });

  // -----------------------------------------------------------------------
  // Edge cases
  // -----------------------------------------------------------------------
  describe("edge cases", () => {
    it("works with maxRetries = 0 (no retries, single attempt)", async () => {
      const noRetryDecorator = new RetryAnalysisDecorator(
        mockPort as unknown as IAnalysisPort,
        0,
        BASE_DELAY_MS,
      );

      mockPort.analyzeChunk.mockResolvedValue(errorResult(0, "only shot"));

      const result = await noRetryDecorator.analyzeChunk(makeChunk(), "general", "es");

      expect(mockPort.analyzeChunk).toHaveBeenCalledOnce();
      expect(result.error).toBe("Chunk 1: only shot");
    });

    it("works with maxRetries = 1 (one retry, two total attempts)", async () => {
      const oneRetryDecorator = new RetryAnalysisDecorator(
        mockPort as unknown as IAnalysisPort,
        1,
        BASE_DELAY_MS,
      );

      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "fail"))
        .mockResolvedValueOnce(successResult(0));

      const promise = oneRetryDecorator.analyzeChunk(makeChunk(), "general", "es");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);

      const result = await promise;

      expect(result).toEqual(successResult(0));
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(2);
    });

    it("preserves suggestions from a successful result unmodified", async () => {
      const suggestions: Suggestion[] = [
        {
          id: "chunk0-0",
          originalText: "mucho muy",
          suggestedText: "muy",
          justification: "Redundancia",
          category: "Redundancia",
          severity: "medium",
        },
        {
          id: "chunk0-1",
          originalText: "en este momento",
          suggestedText: "ahora",
          justification: "Muletilla temporal",
          category: "Muletilla",
          severity: "low",
        },
      ];
      const expected = successResult(0, suggestions);
      mockPort.analyzeChunk.mockResolvedValue(expected);

      const result = await decorator.analyzeChunk(makeChunk(), "general", "es");

      expect(result.suggestions).toEqual(suggestions);
      expect(result.suggestions).toHaveLength(2);
    });

    it("treats result with empty error string as failure and normalizes the message", async () => {
      const result: ChunkResult = { chunkIndex: 0, suggestions: [], error: "" };
      mockPort.analyzeChunk
        .mockResolvedValueOnce(result)
        .mockResolvedValueOnce(result)
        .mockResolvedValueOnce(result)
        .mockResolvedValueOnce(result);

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 4);

      const actual = await promise;

      expect(actual).toEqual({
        chunkIndex: 0,
        suggestions: [],
        error: "Chunk 1: Unknown analysis error",
      });
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(MAX_RETRIES + 1);
    });

    it("handles result with undefined error as success", async () => {
      const result: ChunkResult = { chunkIndex: 0, suggestions: [] };
      mockPort.analyzeChunk.mockResolvedValue(result);

      const actual = await decorator.analyzeChunk(makeChunk(), "general", "es");

      expect(actual).toBe(result);
      expect(mockPort.analyzeChunk).toHaveBeenCalledOnce();
    });

    it("treats whitespace-only error strings as failure and retries", async () => {
      mockPort.analyzeChunk
        .mockResolvedValueOnce(errorResult(0, "   "))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);

      const actual = await promise;

      expect(actual).toEqual(successResult(0));
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(2);
    });

    it("recovers from a thrown exception when a later retry succeeds", async () => {
      mockPort.analyzeChunk
        .mockRejectedValueOnce(new Error("temporary outage"))
        .mockResolvedValueOnce(successResult(0));

      const promise = decorator.analyzeChunk(makeChunk(), "general", "es");
      await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);

      const actual = await promise;

      expect(actual).toEqual(successResult(0));
      expect(mockPort.analyzeChunk).toHaveBeenCalledTimes(2);
    });

    it("preserves legitimate successful results even when suggestions are empty", async () => {
      const result: ChunkResult = { chunkIndex: 0, suggestions: [] };
      mockPort.analyzeChunk.mockResolvedValue(result);

      const actual = await decorator.analyzeChunk(makeChunk(), "general", "es");

      expect(actual).toBe(result);
      expect(actual.error).toBeUndefined();
    });
  });
});
