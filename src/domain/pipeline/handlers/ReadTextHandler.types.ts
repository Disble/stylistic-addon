import type { PipelineContext } from "../PipelineContext";

/** One phase in the document-analysis chain of responsibility. */
export interface PipelineHandler {
  handle(ctx: PipelineContext, next: () => Promise<void>): Promise<void>;
}
