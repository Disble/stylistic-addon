import type { MessageBarProps } from "@fluentui/react-components";
import type { TaskpaneShellStatus } from "../../TaskpaneShellStore.types";

/** Fluent MessageBar intent supported by the status bar. */
export type StatusBarIntent = NonNullable<MessageBarProps["intent"]>;

/** Props required to render the taskpane status bar. */
export type StatusBarProps = Readonly<{
  status: TaskpaneShellStatus;
}>;

/** Griffel class slots consumed by the status bar. */
export type StatusBarClasses = Readonly<{
  root: string;
}>;
