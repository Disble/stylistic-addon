import type { IAnalysisPort, IDocumentPort } from "../domain/ports";
import type { AnalysisProfileId } from "../domain/Profile.types";
import { PipelineOrchestrator } from "../domain/pipeline/PipelineOrchestrator";
import { PipelineStateMachine } from "../domain/pipeline/PipelineStateMachine";
import type { ResultsPanelDeps } from "./SuggestionCardRenderer.types";

/** Mutable cancellation flag shared across taskpane analysis handlers. */
export type TaskpaneCancelState = {
  value: boolean;
};

/** Collaborators required by the taskpane analysis workflow handlers. */
export type TaskpaneAnalysisHandlersRuntime = {
  analysisPort: IAnalysisPort;
  cardRendererDeps: ResultsPanelDeps;
  cancelState: TaskpaneCancelState;
  documentPort: IDocumentPort;
  getSelectedGenero: () => AnalysisProfileId;
  orchestrator: PipelineOrchestrator;
  refreshCleanupVisibility: () => Promise<void>;
  stateMachine: PipelineStateMachine;
};
