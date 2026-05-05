import type { TaskpaneShellProgress } from "../../TaskpaneShellStore";

export type ProgressPanelProps = Readonly<{
  progress: TaskpaneShellProgress;
}>;

export type ProgressPanelClasses = Readonly<{
  root: string;
  message: string;
}>;
