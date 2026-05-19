import type {
  ChunkAnalysisStatus,
  ChunkPollResult,
} from "../../domain/mastra/MastraWorkflow.types";

/** Builds a retryable poll result while preserving the backend run identifier. */
export function buildRetryablePollResult(
  chunkIndex: number,
  runId: string,
  error: string
): ChunkPollResult {
  return {
    chunkIndex,
    runId,
    status: "retryable-failure",
    origin: "frontend-retryable",
    suggestions: [],
    error,
  };
}

/** Returns whether the workflow is still progressing and should be polled again. */
export function isNonTerminalStatus(
  status: string
): status is Extract<ChunkAnalysisStatus, "running" | "pending" | "waiting"> {
  return ["running", "pending", "waiting"].includes(status);
}

/** Returns whether the workflow requires a resume capability the frontend lacks. */
export function requiresResume(status: string): boolean {
  return ["suspended", "paused"].includes(status);
}

/** Returns whether the workflow already reached a terminal backend state. */
export function isTerminalStatus(
  status: string
): status is Exclude<ChunkAnalysisStatus, "running" | "pending" | "waiting"> {
  return ["success", "failed", "tripwire", "canceled", "bailed"].includes(status);
}
