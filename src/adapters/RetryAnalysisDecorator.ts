/* global console, setTimeout */

/**
 * Retry Analysis Decorator — Decorator pattern for `IAnalysisPort`.
 *
 * Adds retry-with-exponential-backoff behavior to any `IAnalysisPort`
 * implementation without modifying the underlying adapter. This separates
 * the retry concern from the Mastra communication concern.
 *
 * The `MastraAdapter` focuses purely on the HTTP protocol. The decorator
 * wraps it transparently, retrying failed `analyzeChunk()` calls up to
 * `maxRetries` times with exponential backoff (`base * 2^attempt` ms).
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
import { TextChunk, ChunkResult } from "../domain/types";

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
   * Analyzes a chunk with up to `maxRetries` retry attempts on failure.
   * Uses exponential backoff: `baseDelayMs * 2^(attempt-1)` between retries.
   * Never throws — always returns a `ChunkResult`.
   */
  async analyzeChunk(chunk: TextChunk, profile: string, language: string): Promise<ChunkResult> {
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
        const result = await this.wrapped.analyzeChunk(chunk, profile, language);
        const errorMessage = this.normalizeError(result.error);

        if (errorMessage === undefined) {
          return result;
        }

        lastError = errorMessage;
        console.warn(
          `⚠️ [RetryDecorator] Chunk #${chunk.index} intento ${attempt} falló: ${lastError}`
        );
      } catch (error) {
        lastError = this.normalizeThrownError(error);
        console.warn(
          `⚠️ [RetryDecorator] Chunk #${chunk.index} intento ${attempt} falló: ${lastError}`
        );
      }
    }

    console.error(
      `❌ [RetryDecorator] Chunk #${chunk.index} agotó ${this.maxRetries} reintentos. Último error: ${lastError}`
    );
    return {
      chunkIndex: chunk.index,
      suggestions: [],
      error: `Chunk ${chunk.index + 1}: ${lastError}`,
    };
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
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
}
