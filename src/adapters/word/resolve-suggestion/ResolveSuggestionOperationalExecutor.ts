import type { ResolutionExecutionReport } from "../../../domain/suggestion/SuggestionResolutionWorkflow.types";
import type {
  ResolutionObservation,
  ResolveSuggestionAction,
} from "./ResolutionContext";
import type { ResolutionErrorSerializer } from "./ResolutionErrorParser";
import type { ResolutionObservabilityReporter } from "./ResolutionObservabilityAdapter";

/**
 * Executes the already-observed wrapper-owned tracked-change collection.
 *
 * This class owns the mutation boundary only. It assumes locating and
 * observation already proved that the wrapper scope is unique and actionable.
 */
export class ResolveSuggestionOperationalExecutor {
  constructor(
    private readonly action: ResolveSuggestionAction,
    private readonly observabilityReporter: ResolutionObservabilityReporter,
    private readonly errorSerializer: ResolutionErrorSerializer,
  ) {}

  /** Applies the requested action to the wrapper-owned tracked-change collection. */
  async execute(
    context: Word.RequestContext,
    observation: ResolutionObservation,
  ): Promise<ResolutionExecutionReport> {
    const collection = observation.trackedChangesCollection;
    if (!collection) {
      return {
        attempted: 0,
        completed: 0,
        remaining: 0,
        error:
          "La sugerencia no expuso una colección operacional ejecutable dentro del wrapper.",
      };
    }

    await this.observabilityReporter.emitPhase("execute", "started", {
      trackedChangesAttempted: observation.trackedChanges.length,
    });

    try {
      if (this.action === "accept") {
        collection.acceptAll();
      } else {
        collection.rejectAll();
      }
      await context.sync();
    } catch (error) {
      const serialized = this.errorSerializer.serialize(error);
      await this.observabilityReporter.emitPhase("execute", "failed", {
        attempted: observation.trackedChanges.length,
        completed: 0,
        remaining: observation.trackedChanges.length,
        error: serialized.message,
      });

      return {
        attempted: observation.trackedChanges.length,
        completed: 0,
        remaining: observation.trackedChanges.length,
        failureIndex: 0,
        error: serialized.message,
      };
    }

    const report: ResolutionExecutionReport = {
      attempted: observation.trackedChanges.length,
      completed: observation.trackedChanges.length,
      remaining: 0,
    };
    await this.observabilityReporter.emitPhase("execute", "succeeded", {
      attempted: report.attempted,
      completed: report.completed,
      remaining: report.remaining,
    });
    return report;
  }
}
