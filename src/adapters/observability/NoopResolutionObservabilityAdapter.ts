import type { IResolutionObservabilityPort } from "../../domain/ports";
import type {
  ResolutionObservabilityEvent,
  ResolutionObservabilitySnapshot,
} from "../../domain/types";

/** Swallows resolution observability records when no concrete adapter is wired. */
export class NoopResolutionObservabilityAdapter
  implements IResolutionObservabilityPort
{
  /** Ignores one phase event without affecting workflow semantics. */
  async emitEvent(_event: ResolutionObservabilityEvent): Promise<void> {
    return undefined;
  }

  /** Ignores one snapshot without affecting workflow semantics. */
  async captureSnapshot(
    _snapshot: ResolutionObservabilitySnapshot,
  ): Promise<void> {
    return undefined;
  }
}
