/** Internal state for one execute-resolution workflow run. */
export type ExecuteResolutionState =
  | "idle"
  | "locating"
  | "observing-before"
  | "executing"
  | "cleaning-comment"
  | "cleaning-anchor"
  | "inspecting-after"
  | "completed"
  | "failed";
