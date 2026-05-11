import type { WorkflowOutput } from "../../domain/mastra/MastraWorkflow.types";
import { MOCK_MASTRA_POLL_OUTPUT } from "./MockMastraPollOutputFactory.constants";

/**
 * Deterministic development output returned when the Mastra poll bypass is
 * enabled from `config.ts`.
 *
 * Keep this payload structurally identical to the backend success contract so
 * the rest of the pipeline behaves exactly as if Mastra had returned it.
 */
/** Returns the deterministic Mastra poll payload used by bypass mode. */
export function createMockMastraPollOutput(): WorkflowOutput {
  return MOCK_MASTRA_POLL_OUTPUT;
}
