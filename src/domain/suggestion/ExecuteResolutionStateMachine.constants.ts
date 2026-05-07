import type { ExecuteResolutionState } from "./ExecuteResolutionStateMachine.types";

/** Valid transitions for one execute-resolution workflow run. */
export const EXECUTE_RESOLUTION_STATE_TRANSITIONS: Record<
  ExecuteResolutionState,
  ExecuteResolutionState[]
> = {
  idle: ["locating", "failed"],
  locating: ["observing-before", "failed", "completed"],
  "observing-before": ["executing", "failed", "completed"],
  executing: ["cleaning-comment", "failed"],
  "cleaning-comment": ["cleaning-anchor", "failed"],
  "cleaning-anchor": ["inspecting-after", "failed"],
  "inspecting-after": ["completed", "failed"],
  completed: [],
  failed: [],
};
