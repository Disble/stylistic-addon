import type { ResolutionExecutionReport } from "../../../domain/types";

/** Applies one terminal resolution action to a tracked-change collection. */
export class TrackedChangeResolutionExecutor {
  constructor(
    private readonly suggestionId: string,
    private readonly action: "accept" | "reject",
  ) {}

  /** Applies the requested action and reports exactly how far execution got. */
  apply(trackedChanges: Word.TrackedChange[]): ResolutionExecutionReport {
    let completed = 0;

    for (const [index, trackedChange] of trackedChanges.entries()) {
      try {
        if (this.action === "accept") {
          trackedChange.accept();
        } else {
          trackedChange.reject();
        }
        completed += 1;
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        console.warn(
          `⚠️ [TrackedChangeResolutionExecutor] ${this.action} failed at index ${index} for suggestionId="${this.suggestionId}": ${message}`,
        );
        return {
          attempted: trackedChanges.length,
          completed,
          remaining: trackedChanges.length - completed,
          failureIndex: index,
          error: message,
        };
      }
    }

    console.log(
      `🎯 [TrackedChangeResolutionExecutor] executed ${this.action} on ${trackedChanges.length} tracked changes for suggestionId="${this.suggestionId}"`,
    );

    return {
      attempted: trackedChanges.length,
      completed,
      remaining: trackedChanges.length - completed,
    };
  }
}
