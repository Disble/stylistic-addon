import type { TaskpaneShellStatus } from "../../TaskpaneShellStore";

export type StatusBarProps = Readonly<{
  status: TaskpaneShellStatus;
}>;

export type StatusBarClasses = Readonly<{
  root: string;
}>;
