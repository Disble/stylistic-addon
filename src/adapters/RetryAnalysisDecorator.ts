/* global console, setTimeout */

/**
 * Retry Analysis Decorator — Decorator pattern for `IAnalysisPort`.
 *
 * Adds retry-with-exponential-backoff behavior to any `IAnalysisPort`
 * implementation without modifying the underlying adapter. This separates
 * the retry concern from the Mastra communication concern.
 *
 * The `MastraAdapter` focuses purely on the HTTP protocol. The decorator
 * wraps it transparently, retrying chunk submission only while no `runId`
 * was confirmed, and retrying chunk polling only on thrown transport errors.
 *
 * Benefits:
 * - `MastraAdapter` is clean and single-purpose.
 * - Tests can inject the raw adapter without retry noise.
 * - A future circuit-breaker or caching layer can be added as another decorator.
 *
 * Usage:
 * ```typescript
 * const analysisPort = new RetryAnalysisDecorator(
 *   new MastraAdapter(),
 *   MAX_RETRIES,
 *   RETRY_BASE_DELAY_MS
 * );
 * ```
 *
 * @module RetryAnalysisDecorator
 */

import { IAnalysisPort } from "../domain/ports";
import { TextChunk, ChunkSubmitResult, ChunkPollResult } from "../domain/types";

export class RetryAnalysisDecorator implements IAnalysisPort {
  constructor(
    private readonly wrapped: IAnalysisPort,
    private readonly maxRetries: number,
    private readonly baseDelayMs: number
  ) {}

  /** Delegates connection check directly — no retry needed. */
  checkConnection(): Promise<boolean> {
    return this.wrapped.checkConnection();
  }

  /**
   * Submits a chunk with up to `maxRetries` retry attempts while no `runId`
   * was obtained. Uses exponential backoff: `baseDelayMs * 2^(attempt-1)`.
   */
  async submitChunkAnalysis(
    chunk: TextChunk,
    profile: string,
    language: string
  ): Promise<ChunkSubmitResult> {
    let lastError = "Unknown analysis error";

    for (let attempt = 0; attempt <= this.maxRetries; attempt++) {
      if (attempt > 0) {
        const delayMs = this.baseDelayMs * Math.pow(2, attempt - 1);
        console.warn(
          `🔄 [RetryDecorator] Retry ${attempt}/${this.maxRetries} chunk #${chunk.index} — esperando ${delayMs}ms`
        );
        await this.delay(delayMs);
      }

      try {
        const result = await this.wrapped.submitChunkAnalysis(chunk, profile, language);

        if (this.hasRunId(result.runId)) {
          return result;
        }

        lastError = this.normalizeError(result.error) ?? "Workflow submission did not return a runId";
        console.warn(
          `⚠️ [RetryDecorator] Submit chunk #${chunk.index} intento ${attempt} falló: ${lastError}`
        );
      } catch (error) {
        lastError = this.normalizeThrownError(error);
        console.warn(
          `⚠️ [RetryDecorator] Submit chunk #${chunk.index} intento ${attempt} falló: ${lastError}`
        );

        if (this.isDeterministicClientError(error)) {
          break;
        }
      }
    }

    console.error(
      `❌ [RetryDecorator] Submit chunk #${chunk.index} agotó ${this.maxRetries} reintentos. Último error: ${lastError}`
    );
    return {
      chunkIndex: chunk.index,
      error: `Chunk ${chunk.index + 1}: ${lastError}`,
    };
  }

  /**
   * Polls a submitted chunk run. Retries only on thrown transport errors;
   * intermediate statuses like `running` are returned as-is.
   */
  async pollChunkAnalysis(chunkIndex: number, runId: string): Promise<ChunkPollResult> {
    let lastError = "Unknown analysis error";

    for (let attempt = 0; attempt <= this.maxRetries; attempt++) {
      if (attempt > 0) {
        const delayMs = this.baseDelayMs * Math.pow(2, attempt - 1);
        console.warn(
          `🔄 [RetryDecorator] Retry ${attempt}/${this.maxRetries} poll chunk #${chunkIndex} — esperando ${delayMs}ms`
        );
        await this.delay(delayMs);
      }

      try {
        return await this.wrapped.pollChunkAnalysis(chunkIndex, runId);
      } catch (error) {
        lastError = this.normalizeThrownError(error);
        console.warn(
          `⚠️ [RetryDecorator] Poll chunk #${chunkIndex} intento ${attempt} falló: ${lastError}`
        );

        if (this.isDeterministicClientError(error)) {
          break;
        }
      }
    }

    console.error(
      `❌ [RetryDecorator] Poll chunk #${chunkIndex} agotó ${this.maxRetries} reintentos. Último error: ${lastError}`
    );
    return {
      chunkIndex,
      runId,
      status: "failed",
      suggestions: [],
      error: `Chunk ${chunkIndex + 1}: ${lastError}`,
    };
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private hasRunId(runId: string | undefined): runId is string {
    return typeof runId === "string" && runId.trim().length > 0;
  }

  private normalizeError(error: string | undefined): string | undefined {
    if (error === undefined) {
      return undefined;
    }

    const normalized = error.trim();
    return normalized.length > 0 ? normalized : "Unknown analysis error";
  }

  private normalizeThrownError(error: unknown): string {
    if (error instanceof Error) {
      const normalized = error.message.trim();
      return normalized.length > 0 ? normalized : "Unknown analysis error";
    }

    if (typeof error === "string") {
      const normalized = error.trim();
      return normalized.length > 0 ? normalized : "Unknown analysis error";
    }

    return "Unknown analysis error";
  }

  private isDeterministicClientError(error: unknown): boolean {
    const status = this.readStatusCode(error);

    if (status === undefined) {
      return false;
    }

    return status >= 400 && status < 500 && status !== 408 && status !== 429;
  }

  private readStatusCode(error: unknown): number | undefined {
    if (typeof error !== "object" || error === null || !("status" in error)) {
      return undefined;
    }

    const { status } = error as { status?: unknown };
    return typeof status === "number" && Number.isInteger(status) ? status : undefined;
  }
}
