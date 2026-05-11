import type { PipelineState } from "./PipelineStateMachine.types";

/** Valid transitions for the analysis pipeline lifecycle. */
export const PIPELINE_STATE_TRANSITIONS: Record<PipelineState, PipelineState[]> = {
  idle: ["reading"],
  reading: ["connecting", "idle", "error"],
  connecting: ["chunking", "idle", "error"],
  chunking: ["analyzing", "idle", "error"],
  analyzing: ["applying", "idle", "error"],
  applying: ["done", "idle", "error"],
  done: ["idle"],
  error: ["idle"],
};
