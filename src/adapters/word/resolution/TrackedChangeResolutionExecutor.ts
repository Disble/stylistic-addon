import type { ResolutionExecutionReport } from "../../../domain/types";
import type { ReplaceResolutionStrategy } from "./ReplaceResolutionStrategyContext";

type BodyTrackedChangeCountProbe =
  | {
      status: "known";
      count: number;
    }
  | {
      status: "unknown";
      error: string;
    };

/** Applies one terminal resolution action to a tracked-change collection. */
export class TrackedChangeResolutionExecutor {
  constructor(
    private readonly suggestionId: string,
    private readonly action: "accept" | "reject",
    private readonly replaceResolutionStrategy: ReplaceResolutionStrategy,
  ) {}

  /** Builds one stable tracked-change diagnostic entry for runtime logs. */
  private describeTrackedChange(trackedChange: Word.TrackedChange): {
    type: string;
  } {
    return {
      type: trackedChange.type ?? "unknown",
    };
  }

  /** Builds a compact tracked-change list so one host attempt can be reconstructed later. */
  private describeTrackedChanges(trackedChanges: Word.TrackedChange[]): Array<{
    type: string;
  }> {
    return trackedChanges.map((trackedChange) =>
      this.describeTrackedChange(trackedChange),
    );
  }

  /** Returns true when the tracked-change type can be semantically verified by body count. */
  private isVerifiableTrackedChangeType(
    trackedChangeType: string,
  ): trackedChangeType is "Added" | "Deleted" {
    return trackedChangeType === "Added" || trackedChangeType === "Deleted";
  }

  /** Returns the known body count from a probe, if Word exposed one. */
  private getKnownBodyTrackedChangeCount(
    probe: BodyTrackedChangeCountProbe,
  ): number | undefined {
    return probe.status === "known" ? probe.count : undefined;
  }

  /** Builds an unverified-mutation signal without conflating unknown host state with count zero. */
  private buildUnverifiedMutation(
    stepIndex: number,
    trackedChangeType: "Added" | "Deleted",
    beforeProbe: BodyTrackedChangeCountProbe,
    afterProbe: BodyTrackedChangeCountProbe,
  ): ResolutionExecutionReport["unverifiedMutation"] | undefined {
    if (beforeProbe.status === "known" && afterProbe.status === "known") {
      return undefined;
    }

    return {
      stepIndex,
      trackedChangeType,
      reason: "body-count-probe-failed",
      ...(beforeProbe.status === "known"
        ? { bodyTrackedChangeCountBefore: beforeProbe.count }
        : { bodyTrackedChangeCountBeforeError: beforeProbe.error }),
      ...(afterProbe.status === "known"
        ? { bodyTrackedChangeCountAfter: afterProbe.count }
        : { bodyTrackedChangeCountAfterError: afterProbe.error }),
    };
  }

  /** Orders tracked changes using the shared replace policy for this action. */
  private orderTrackedChangesForExecution(
    trackedChanges: Word.TrackedChange[],
  ): Word.TrackedChange[] {
    return trackedChanges
      .map((trackedChange, index) => ({
        trackedChange,
        index,
        priority: this.replaceResolutionStrategy.priorityFor(
          trackedChange.type ?? "unknown",
        ),
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
    const orderedTrackedChanges =
      this.orderTrackedChangesForExecution(trackedChanges);
    console.log(
      `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} workflow-plan orderedTypes=${orderedTrackedChanges
        .map((trackedChange) => trackedChange.type ?? "unknown")
        .join(",")}`,
    );
    console.log(
      `🧾 [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} workflow-plan`,
      {
        workflowPlan: {
          inputTrackedChanges: this.describeTrackedChanges(trackedChanges),
          orderedTrackedChanges: this.describeTrackedChanges(
            orderedTrackedChanges,
          ),
        },
      },
    );
    let completed = 0;
    let silentNoOpDetected: ResolutionExecutionReport["silentNoOpDetected"];
    let unverifiedMutation: ResolutionExecutionReport["unverifiedMutation"];

    for (const [index, trackedChange] of orderedTrackedChanges.entries()) {
      const stepResult = await this.applyTrackedChangeStep(
        context,
        trackedChange,
        index,
      );
      completed += stepResult.completed;

      if (!unverifiedMutation && stepResult.unverifiedMutation) {
        unverifiedMutation = stepResult.unverifiedMutation;
      }

      if (!silentNoOpDetected && stepResult.silentNoOpDetected) {
        silentNoOpDetected = stepResult.silentNoOpDetected;
      }

      if (stepResult.error) {
        return {
          attempted: orderedTrackedChanges.length,
          completed,
          remaining: orderedTrackedChanges.length - completed,
          failureIndex: index,
          error: stepResult.error,
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
      ...(silentNoOpDetected ? { silentNoOpDetected } : {}),
      ...(unverifiedMutation ? { unverifiedMutation } : {}),
    };
  }

  /** Applies one tracked-change step and returns diagnostics without mutating the outer loop state. */
  private async applyTrackedChangeStep(
    context: Word.RequestContext,
    trackedChange: Word.TrackedChange,
    index: number,
  ): Promise<{
    completed: number;
    error?: string;
    silentNoOpDetected?: ResolutionExecutionReport["silentNoOpDetected"];
    unverifiedMutation?: ResolutionExecutionReport["unverifiedMutation"];
  }> {
    let actionQueued = false;
    const trackedChangeDetail = this.describeTrackedChange(trackedChange);
    const bodyTrackedChangeCountBefore =
      await this.countBodyTrackedChanges(context);

    try {
      console.log(
        `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} step=${index} type=${trackedChangeDetail.type} stage=queue`,
      );
      if (this.action === "accept") {
        trackedChange.accept();
      } else {
        trackedChange.reject();
      }

      actionQueued = true;
      console.log(
        `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} step=${index} type=${trackedChangeDetail.type} stage=sync-start`,
      );
      await context.sync();
      console.log(
        `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} step=${index} type=${trackedChangeDetail.type} stage=sync-done`,
      );

      const bodyTrackedChangeCountAfter =
        await this.countBodyTrackedChanges(context);

      return {
        completed: 1,
        ...this.buildStepDiagnostics(
          index,
          trackedChangeDetail.type,
          bodyTrackedChangeCountBefore,
          bodyTrackedChangeCountAfter,
        ),
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [TrackedChangeResolutionExecutor] ${this.action} failed at index ${index} for suggestionId="${this.suggestionId}": ${message}`,
      );
      return {
        completed: actionQueued ? 1 : 0,
        error: message,
      };
    }
  }

  /** Computes optional verification diagnostics for one executed tracked-change step. */
  private buildStepDiagnostics(
    index: number,
    trackedChangeType: string,
    beforeProbe: BodyTrackedChangeCountProbe,
    afterProbe: BodyTrackedChangeCountProbe,
  ): {
    silentNoOpDetected?: ResolutionExecutionReport["silentNoOpDetected"];
    unverifiedMutation?: ResolutionExecutionReport["unverifiedMutation"];
  } {
    const beforeCount = this.getKnownBodyTrackedChangeCount(beforeProbe);
    const afterCount = this.getKnownBodyTrackedChangeCount(afterProbe);

    if (!this.isVerifiableTrackedChangeType(trackedChangeType)) {
      return {};
    }

    const unverifiedMutation = this.buildUnverifiedMutation(
      index,
      trackedChangeType,
      beforeProbe,
      afterProbe,
    );
    if (unverifiedMutation) {
      console.warn(
        `⚠️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} step=${index} type=${trackedChangeType} mutation verification unavailable: bodyTrackedChangeCount before=${beforeCount ?? "unknown"} after=${afterCount ?? "unknown"}`,
        unverifiedMutation,
      );
    }

    const silentNoOpDetected =
      beforeCount !== undefined &&
      afterCount !== undefined &&
      beforeCount > 0 &&
      afterCount >= beforeCount
        ? {
            stepIndex: index,
            trackedChangeType,
            bodyTrackedChangeCountBefore: beforeCount,
            bodyTrackedChangeCountAfter: afterCount,
          }
        : undefined;

    if (silentNoOpDetected) {
      console.warn(
        `⚠️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} step=${index} type=${trackedChangeType} silent-no-op detected: bodyTrackedChangeCount before=${beforeCount} after=${afterCount} (proxy mutation did not reduce document tracked-change count)`,
      );
    }

    return {
      ...(unverifiedMutation ? { unverifiedMutation } : {}),
      ...(silentNoOpDetected ? { silentNoOpDetected } : {}),
    };
  }

  /** Applies all tracked changes in one host batch, matching the pre-refactor replace behavior. */
  async applyAtomically(
    context: Word.RequestContext,
    trackedChanges: Word.TrackedChange[],
  ): Promise<ResolutionExecutionReport> {
    const orderedTrackedChanges =
      this.orderTrackedChangesForExecution(trackedChanges);
    console.log(
      `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} workflow-plan atomic-orderedTypes=${orderedTrackedChanges
        .map((trackedChange) => trackedChange.type ?? "unknown")
        .join(",")}`,
    );
    console.log(
      `🧾 [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} workflow-plan atomic`,
      {
        workflowPlan: {
          inputTrackedChanges: this.describeTrackedChanges(trackedChanges),
          orderedTrackedChanges: this.describeTrackedChanges(
            orderedTrackedChanges,
          ),
        },
      },
    );

    try {
      for (const [index, trackedChange] of orderedTrackedChanges.entries()) {
        const trackedChangeDetail = this.describeTrackedChange(trackedChange);
        console.log(
          `⚙️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" action=${this.action} atomic-step=${index} type=${trackedChangeDetail.type} stage=queue`,
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

  /**
   * Counts how many tracked changes the document body currently exposes.
   *
   * Used by `apply()` to detect silent no-op resolutions: when the host
   * accepts/rejects a stale `Word.TrackedChange` proxy (typical for
   * `ccRange.getTrackedChanges()` when the deletion mark lives outside the
   * suggestion CC range), `context.sync()` resolves cleanly but the document
   * tracked-change count does not decrease. Comparing the count before and
   * after each step lets the executor surface this case to the outer command
   * so it can recover with a fresh proxy from a different evidence source.
   *
   * Returns an explicit unknown probe when the host does not expose the count.
   * This must not abort the mutation flow, but callers also must not interpret
   * the unknown state as an actual zero-count document.
   */
  private async countBodyTrackedChanges(
    context: Word.RequestContext,
  ): Promise<BodyTrackedChangeCountProbe> {
    try {
      const bodyTrackedChanges = context.document.body.getTrackedChanges();
      bodyTrackedChanges.load({ select: "type" });
      await context.sync();
      return {
        status: "known",
        count: bodyTrackedChanges.items.length,
      };
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(
        `⚠️ [TrackedChangeResolutionExecutor] suggestionId="${this.suggestionId}" body tracked-change count probe failed: ${message}`,
      );
      return {
        status: "unknown",
        error: message,
      };
    }
  }
}
