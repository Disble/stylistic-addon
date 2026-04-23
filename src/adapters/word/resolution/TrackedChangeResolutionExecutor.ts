import type { ResolutionExecutionReport } from "../../../domain/types";

/** Applies one terminal resolution action to a tracked-change collection. */
export class TrackedChangeResolutionExecutor {
  constructor(
    private readonly suggestionId: string,
    private readonly action: "accept" | "reject",
  ) {}

  /** Builds one stable tracked-change diagnostic entry for runtime logs. */
  private describeTrackedChange(
    trackedChange: Word.TrackedChange,
  ): {
    id: string;
    type: string;
  } {
    return {
      id: String((trackedChange as { id?: string | number }).id ?? "no-id"),
      type: trackedChange.type ?? "unknown",
    };
  }

  /** Builds a compact tracked-change list so one host attempt can be reconstructed later. */
  private describeTrackedChanges(
    trackedChanges: Word.TrackedChange[],
  ): Array<{
    id: string;
    type: string;
  }> {
    return trackedChanges.map((trackedChange) =>
      this.describeTrackedChange(trackedChange),
    );
  }

  /**
   * Orders replace-pair tracked changes so semantic sides resolve predictably.
   *
   * For BOTH accept and reject we process the Deleted side first, then the
   * Added side. The suggestion CC wraps the Added (inserted) text, so
   * resolving the Added TC first destroys the CC and breaks any re-observation
   * the outer command does between steps. Deleted-first keeps the CC anchor
   * stable for the second step in both actions.
   */
  private orderTrackedChangesForExecution(
    trackedChanges: Word.TrackedChange[],
  ): Word.TrackedChange[] {
    const getPriority = (trackedChange: Word.TrackedChange): number => {
      if (trackedChange.type === "Deleted") {
        return 0;
      }

      if (trackedChange.type === "Added") {
        return 1;
      }

      return 2;
    };

    return trackedChanges
      .map((trackedChange, index) => ({
        trackedChange,
        index,
        priority: getPriority(trackedChange),
      }))
      .sort((left, right) => {
        if (left.priority !== right.priority) {
          return left.priority - right.priority;
        }

        return left.index - right.index;
      })
      .map(({ trackedChange }) => trackedChange);
  }

  /** Applies the requested action and reports exactly how far execution got. */
  async apply(
    context: Word.RequestContext,
    trackedChanges: Word.TrackedChange[],
  ): Promise<ResolutionExecutionReport> {
    const orderedTrackedChanges = this.orderTrackedChangesForExecution(
      trackedChanges,
    );
    console.log(
      `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} orderedTypes=${orderedTrackedChanges
        .map((trackedChange) => trackedChange.type ?? "unknown")
        .join(",")}`,
    );
    console.log(
      `🧾 [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} apply-detail`,
      {
        input: this.describeTrackedChanges(trackedChanges),
        ordered: this.describeTrackedChanges(orderedTrackedChanges),
      },
    );
    let completed = 0;

    for (const [index, trackedChange] of orderedTrackedChanges.entries()) {
      let actionQueued = false;
      const trackedChangeDetail = this.describeTrackedChange(trackedChange);

      try {
        console.log(
          `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} step=${index} id=${trackedChangeDetail.id} type=${trackedChangeDetail.type} stage=queue`,
        );
        if (this.action === "accept") {
          trackedChange.accept();
        } else {
          trackedChange.reject();
        }

        actionQueued = true;
        console.log(
          `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} step=${index} id=${trackedChangeDetail.id} type=${trackedChangeDetail.type} stage=sync-start`,
        );
        await context.sync();
        completed += 1;
        console.log(
          `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} step=${index} id=${trackedChangeDetail.id} type=${trackedChangeDetail.type} stage=sync-done`,
        );
      } catch (error) {
        if (actionQueued) {
          completed += 1;
        }

        const message = error instanceof Error ? error.message : String(error);
        console.warn(
          `⚠️ [TrackedChangeResolutionExecutor] ${this.action} failed at index ${index} for suggestionId="${this.suggestionId}": ${message}`,
        );
        return {
          attempted: orderedTrackedChanges.length,
          completed,
          remaining: orderedTrackedChanges.length - completed,
          failureIndex: index,
          error: message,
        };
      }
    }

    console.log(
      `🎯 [TrackedChangeResolutionExecutor] executed ${this.action} on ${orderedTrackedChanges.length} tracked changes for suggestionId="${this.suggestionId}"`,
    );

    return {
      attempted: orderedTrackedChanges.length,
      completed,
      remaining: orderedTrackedChanges.length - completed,
    };
  }

  /** Applies all tracked changes in one host batch, matching the pre-refactor replace behavior. */
  async applyAtomically(
    context: Word.RequestContext,
    trackedChanges: Word.TrackedChange[],
  ): Promise<ResolutionExecutionReport> {
    const orderedTrackedChanges = this.orderTrackedChangesForExecution(
      trackedChanges,
    );
    console.log(
      `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} atomic-orderedTypes=${orderedTrackedChanges
        .map((trackedChange) => trackedChange.type ?? "unknown")
        .join(",")}`,
    );
    console.log(
      `🧾 [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} atomic-detail`,
      {
        input: this.describeTrackedChanges(trackedChanges),
        ordered: this.describeTrackedChanges(orderedTrackedChanges),
      },
    );

    try {
      for (const [index, trackedChange] of orderedTrackedChanges.entries()) {
        const trackedChangeDetail = this.describeTrackedChange(trackedChange);
        console.log(
          `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} atomic-step=${index} id=${trackedChangeDetail.id} type=${trackedChangeDetail.type} stage=queue`,
        );
        if (this.action === "accept") {
          trackedChange.accept();
        } else {
          trackedChange.reject();
        }
      }

      console.log(
        `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} stage=atomic-sync-start`,
      );
      await context.sync();
      console.log(
        `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} stage=atomic-sync-done`,
      );
      console.log(
        `🎯 [TrackedChangeResolutionExecutor] executed ${this.action} atomically on ${orderedTrackedChanges.length} tracked changes for suggestionId="${this.suggestionId}"`,
      );

      return {
        attempted: orderedTrackedChanges.length,
        completed: orderedTrackedChanges.length,
        remaining: 0,
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [TrackedChangeResolutionExecutor] atomic ${this.action} failed for suggestionId="${this.suggestionId}": ${message}`,
      );
      return {
        attempted: orderedTrackedChanges.length,
        completed: 0,
        remaining: orderedTrackedChanges.length,
        failureIndex: 0,
        error: message,
      };
    }
  }
}
