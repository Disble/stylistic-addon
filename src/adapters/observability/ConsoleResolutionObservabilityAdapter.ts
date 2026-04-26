/* global console */

import type { IResolutionObservabilityPort } from "../../domain/ports";
import type {
  ResolutionObservabilityEvent,
  ResolutionObservabilitySnapshot,
} from "../../domain/suggestion/SuggestionResolutionWorkflow.types";

/** Emits best-effort structured resolution observability to the console. */
export class ConsoleResolutionObservabilityAdapter
  implements IResolutionObservabilityPort
{
  /** Writes one structured phase event without affecting workflow semantics. */
  async emitEvent(event: ResolutionObservabilityEvent): Promise<void> {
    console.log("📡 [ResolutionEvent]", event);
  }

  /** Writes one structured host-evidence snapshot without affecting workflow semantics. */
  async captureSnapshot(
    snapshot: ResolutionObservabilitySnapshot,
  ): Promise<void> {
    console.log("🧪 [ResolutionSnapshot]", snapshot);
  }
}
