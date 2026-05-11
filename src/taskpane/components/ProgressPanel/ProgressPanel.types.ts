import type { TaskpaneShellProgress } from "../../TaskpaneShellStore.types";

/** Props required to render the compact progress panel. */
export type ProgressPanelProps = Readonly<{
  progress: TaskpaneShellProgress;
  onCancelAnalysis: () => Promise<void> | void;
  onRetryAnalysisQuery: () => Promise<void> | void;
}>;

/** Griffel class slots consumed by the compact progress panel. */
export type ProgressPanelClasses = Readonly<{
  root: string;
  header: string;
  message: string;
  actions: string;
}>;
