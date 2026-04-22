/* global console */

import type { ITelemetryPort } from "../../domain/ports";
import type { ResolutionTelemetryEvent } from "../../domain/types";

/** Emits best-effort structured resolution telemetry to the console. */
export class ConsoleTelemetryAdapter implements ITelemetryPort {
  /** Writes one structured event without affecting workflow semantics. */
  async emit(event: ResolutionTelemetryEvent): Promise<void> {
    console.log("📡 [ResolutionTelemetry]", event);
  }
}
