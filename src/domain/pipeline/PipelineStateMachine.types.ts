/** All possible states of the analysis pipeline. */
export type PipelineState =
  | "idle"
  | "reading"
  | "connecting"
  | "chunking"
  | "analyzing"
  | "applying"
  | "done"
  | "error";
