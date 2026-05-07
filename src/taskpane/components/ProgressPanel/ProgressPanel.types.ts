import type { TaskpaneShellProgress } from "../../TaskpaneShellStore.types";

/** Props required to render pipeline progress feedback. */
export type ProgressPanelProps = Readonly<{
  progress: TaskpaneShellProgress;
}>;

/** Griffel class slots consumed by the progress panel. */
export type ProgressPanelClasses = Readonly<{
  root: string;
  message: string;
}>;
