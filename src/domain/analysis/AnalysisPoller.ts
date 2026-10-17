/* global setTimeout */

import type { IAnalysisPort } from "../ports";
import type { AnalysisSnapshot } from "./Analysis.types";
import type { AnalysisPollOptions, AnalysisPollOutcome } from "./AnalysisPoller.types";

/**
 * Polls one analysis until it leaves `running`, the caller cancels, or a read
 * fails. Shared by the analysis pipeline and the taskpane retry-query action so
 * both wait for an analysis the same way.
 */
export class AnalysisPoller {
  constructor(
    private readonly analysisPort: IAnalysisPort,
    private readonly intervalMs: number
  ) {}

  /** Waits for `current` to finish, reading it again every `intervalMs`. */
  async waitUntilFinished(
    current: AnalysisSnapshot,
    options: AnalysisPollOptions = {}
  ): Promise<AnalysisPollOutcome> {
    let snapshot = current;

    while (snapshot.status === "running") {
      if (options.shouldCancel?.()) {
        return { kind: "canceled-locally" };
      }

      await this.delay(this.intervalMs);

      try {
        snapshot = await this.analysisPort.getAnalysis(current.analysisId);
      } catch (error) {
        return { kind: "poll-failed", message: this.readErrorMessage(error) };
      }

      options.onPoll?.(snapshot);
    }

    return { kind: "finished", snapshot };
  }

  /** Reads a displayable message from a thrown value. */
  private readErrorMessage(error: unknown): string {
    if (error instanceof Error && error.message.trim().length > 0) {
      return error.message.trim();
    }

    return "error desconocido";
  }

  /** Waits between reads to avoid hammering the backend. */
  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
