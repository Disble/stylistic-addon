import type { WorkflowGenero } from "./mastra/MastraWorkflow.types";

/** Canonical identifier shared by persisted analysis profiles and workflow genre input. */
export type AnalysisProfileId = WorkflowGenero;

/** An analysis profile option shown in the UI dropdown. */
export interface Profile {
  /** Machine-readable identifier sent to the workflow. */
  id: AnalysisProfileId;

  /** Human-readable label displayed in the dropdown. */
  label: string;
}
