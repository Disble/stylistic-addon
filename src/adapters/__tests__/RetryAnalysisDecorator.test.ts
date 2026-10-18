import type { IAnalysisPort } from "../../domain/ports";
import type { AnalysisRequest, AnalysisSnapshot } from "../../domain/analysis/Analysis.types";
import { AnalysisRequestError } from "../../domain/analysis/AnalysisRequestError";
import { RetryAnalysisDecorator } from "../RetryAnalysisDecorator";

type MockAnalysisPort = IAnalysisPort & {
  getReadiness: ReturnType<typeof vi.fn<IAnalysisPort["getReadiness"]>>;
  startAnalysis: ReturnType<typeof vi.fn<IAnalysisPort["startAnalysis"]>>;
  getAnalysis: ReturnType<typeof vi.fn<IAnalysisPort["getAnalysis"]>>;
  cancelAnalysis: ReturnType<typeof vi.fn<IAnalysisPort["cancelAnalysis"]>>;
};

/** Creates a mock IAnalysisPort with vi.fn() stubs. */
function createMockPort(): MockAnalysisPort {
  return {
    getReadiness: vi.fn<IAnalysisPort["getReadiness"]>(),
    startAnalysis: vi.fn<IAnalysisPort["startAnalysis"]>(),
    getAnalysis: vi.fn<IAnalysisPort["getAnalysis"]>(),
    cancelAnalysis: vi.fn<IAnalysisPort["cancelAnalysis"]>(),
  };
}

function makeRequest(): AnalysisRequest {
  return {
    documentUuid: "11111111-1111-4111-8111-111111111111",
    genre: "general",
    chunks: [{ text: "Some text to analyze.", index: 0, total: 1, startOffset: 0 }],
  };
}

function makeSnapshot(status: AnalysisSnapshot["status"] = "running"): AnalysisSnapshot {
  return { analysisId: "a1", status, chunks: [] };
}

describe("RetryAnalysisDecorator", () => {
  const MAX_RETRIES = 3;
  const BASE_DELAY_MS = 1_000;

  let mockPort: MockAnalysisPort;
  let decorator: RetryAnalysisDecorator;

  beforeEach(() => {
    vi.useFakeTimers();
    mockPort = createMockPort();
    decorator = new RetryAnalysisDecorator(mockPort, MAX_RETRIES, BASE_DELAY_MS);
    vi.spyOn(console, "warn").mockImplementation(() => {});
    vi.spyOn(console, "error").mockImplementation(() => {});
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
  });

  it("delegates getReadiness without retrying", async () => {
    mockPort.getReadiness.mockResolvedValue({ ready: false, reason: "No model" });

    await expect(decorator.getReadiness()).resolves.toEqual({ ready: false, reason: "No model" });
    expect(mockPort.getReadiness).toHaveBeenCalledOnce();
  });

  it("returns the started analysis on the first successful attempt", async () => {
    mockPort.startAnalysis.mockResolvedValue(makeSnapshot());

    await expect(decorator.startAnalysis(makeRequest())).resolves.toEqual(makeSnapshot());
    expect(mockPort.startAnalysis).toHaveBeenCalledOnce();
    expect(mockPort.startAnalysis).toHaveBeenCalledWith(makeRequest());
  });

  it("retries a transport failure with exponential backoff, then succeeds", async () => {
    mockPort.startAnalysis
      .mockRejectedValueOnce(new AnalysisRequestError("Failed to fetch"))
      .mockRejectedValueOnce(new AnalysisRequestError("Service Unavailable", 503))
      .mockResolvedValueOnce(makeSnapshot());

    const promise = decorator.startAnalysis(makeRequest());
    await vi.advanceTimersByTimeAsync(BASE_DELAY_MS);
    expect(mockPort.startAnalysis).toHaveBeenCalledTimes(2);
    await vi.advanceTimersByTimeAsync(BASE_DELAY_MS * 2);

    await expect(promise).resolves.toEqual(makeSnapshot());
    expect(mockPort.startAnalysis).toHaveBeenCalledTimes(3);
  });

  it("throws the last error after exhausting retries", async () => {
    mockPort.startAnalysis.mockRejectedValue(new AnalysisRequestError("Failed to fetch"));

    const settled = decorator.startAnalysis(makeRequest()).catch((error: unknown) => error);
    await vi.runAllTimersAsync();

    expect(await settled).toEqual(new AnalysisRequestError("Failed to fetch"));
    expect(mockPort.startAnalysis).toHaveBeenCalledTimes(MAX_RETRIES + 1);
  });

  it.each([400, 401, 404, 409])("does not retry a deterministic %i refusal", async (status) => {
    mockPort.startAnalysis.mockRejectedValue(new AnalysisRequestError("Refused", status));

    await expect(decorator.startAnalysis(makeRequest())).rejects.toThrow("Refused");
    expect(mockPort.startAnalysis).toHaveBeenCalledOnce();
  });

  it.each([408, 429])("retries a %i answer", async (status) => {
    mockPort.startAnalysis
      .mockRejectedValueOnce(new AnalysisRequestError("Slow down", status))
      .mockResolvedValueOnce(makeSnapshot());

    const promise = decorator.startAnalysis(makeRequest());
    await vi.runAllTimersAsync();

    await expect(promise).resolves.toEqual(makeSnapshot());
    expect(mockPort.startAnalysis).toHaveBeenCalledTimes(2);
  });

  it("retries getAnalysis on transport failures and returns the snapshot", async () => {
    mockPort.getAnalysis
      .mockRejectedValueOnce(new AnalysisRequestError("Failed to fetch"))
      .mockResolvedValueOnce(makeSnapshot("completed"));

    const promise = decorator.getAnalysis("a1");
    await vi.runAllTimersAsync();

    await expect(promise).resolves.toEqual(makeSnapshot("completed"));
    expect(mockPort.getAnalysis).toHaveBeenCalledTimes(2);
    expect(mockPort.getAnalysis).toHaveBeenCalledWith("a1");
  });

  it("throws from getAnalysis after exhausting retries so the caller can offer a retry query", async () => {
    mockPort.getAnalysis.mockRejectedValue(new AnalysisRequestError("Failed to fetch"));

    const settled = decorator.getAnalysis("a1").catch((error: unknown) => error);
    await vi.runAllTimersAsync();

    expect(await settled).toEqual(new AnalysisRequestError("Failed to fetch"));
    expect(mockPort.getAnalysis).toHaveBeenCalledTimes(MAX_RETRIES + 1);
  });

  it("delegates cancelAnalysis without retrying", async () => {
    mockPort.cancelAnalysis.mockRejectedValue(new AnalysisRequestError("Failed to fetch"));

    await expect(decorator.cancelAnalysis("a1")).rejects.toThrow("Failed to fetch");
    expect(mockPort.cancelAnalysis).toHaveBeenCalledOnce();
    expect(mockPort.cancelAnalysis).toHaveBeenCalledWith("a1");
  });
});
