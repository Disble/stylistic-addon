import type { ChunkRunReference } from "../../mastra/MastraWorkflow.types";
import type { Suggestion } from "../../suggestion/Suggestion.types";
import type { PipelineContext } from "../PipelineContext";

/** One text chunk submitted to the backend analysis workflow. */
export type AnalyzeChunk = NonNullable<PipelineContext["chunks"]>[number];

/** Mutable state shared between chunk submission and polling phases. */
export type AnalysisState = {
  scope: string;
  chunks: NonNullable<PipelineContext["chunks"]>;
  allSuggestions: Suggestion[];
  chunkErrors: string[];
  pendingRuns: Map<number, string>;
  activeRunReferences: Map<number, ChunkRunReference>;
  retryableRunReferences: Map<number, ChunkRunReference>;
  totalSteps: number;
  submittedCount: number;
  completedCount: number;
};
