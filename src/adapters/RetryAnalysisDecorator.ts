/* global console, setTimeout */

/**
 * Retry Analysis Decorator — Decorator pattern for `IAnalysisPort`.
 *
 * Adds retry-with-exponential-backoff behavior to any `IAnalysisPort`
 * implementation without modifying the underlying adapter. This separates
 * the retry concern from the HTTP communication concern.
 *
 * Starting and reading an analysis are retried on transport failures, 5xx,
 * 408 and 429. Any other 4xx is a deterministic refusal and is thrown at once,
 * with the backend's explanation intact. Readiness and cancellation are
 * delegated as-is: readiness never throws, and cancel stays user-driven.
 *
 * Usage:
 * ```typescript
 * const analysisPort = new RetryAnalysisDecorator(
 *   new HubAnalysisAdapter(httpClient),
 *   MAX_RETRIES,
 *   RETRY_BASE_DELAY_MS
 * );
 * ```
 *
 * @module RetryAnalysisDecorator
 */

import type {
  AnalysisReadiness,
  AnalysisRequest,
  AnalysisSnapshot,
} from "../domain/analysis/Analysis.types";
import type { IAnalysisPort } from "../domain/ports";

/** Adds retry and backoff semantics to an analysis port. */
export class RetryAnalysisDecorator implements IAnalysisPort {
  constructor(
    private readonly wrapped: IAnalysisPort,
    private readonly maxRetries: number,
    private readonly baseDelayMs: number
  ) {}

  /** Delegates the readiness check directly — no retry needed. */
  getReadiness(): Promise<AnalysisReadiness> {
    return this.wrapped.getReadiness();
  }

  /** Starts an analysis, retrying transient failures. */
  startAnalysis(request: AnalysisRequest): Promise<AnalysisSnapshot> {
    return this.withRetry("start analysis", () => this.wrapped.startAnalysis(request));
  }

  /** Reads an analysis, retrying transient failures. */
  getAnalysis(analysisId: string): Promise<AnalysisSnapshot> {
    return this.withRetry(`poll analysis ${analysisId}`, () =>
      this.wrapped.getAnalysis(analysisId)
    );
  }

  /** Delegates cancellation directly — retry semantics stay user-driven. */
  cancelAnalysis(analysisId: string): Promise<void> {
    return this.wrapped.cancelAnalysis(analysisId);
  }

  /** Runs `operation` up to `maxRetries + 1` times with exponential backoff. */
  private async withRetry<T>(label: string, operation: () => Promise<T>): Promise<T> {
    let lastError: unknown;

    for (let attempt = 0; attempt <= this.maxRetries; attempt++) {
      if (attempt > 0) {
        const delayMs = this.baseDelayMs * 2 ** (attempt - 1);
        console.warn(
          `🔄 [RetryDecorator] Retry ${attempt}/${this.maxRetries} ${label} — esperando ${delayMs}ms`
        );
        await this.delay(delayMs);
      }

      try {
        return await operation();
      } catch (error) {
        lastError = error;
        console.warn(`⚠️ [RetryDecorator] ${label} intento ${attempt} falló:`, error);

        if (this.isDeterministicClientError(error)) {
          throw error;
        }
      }
    }

    console.error(`❌ [RetryDecorator] ${label} agotó ${this.maxRetries} reintentos.`);
    throw lastError;
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
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
