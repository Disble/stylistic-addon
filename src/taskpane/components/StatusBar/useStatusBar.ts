import { makeStyles, tokens } from "@fluentui/react-components";
import type { MessageBarProps } from "@fluentui/react-components";
import type { StatusBarClasses } from "./StatusBar.types";
import type { TaskpaneStatusType } from "../../TaskpaneShellStore";

type MessageBarIntent = NonNullable<MessageBarProps["intent"]>;

const STATUS_INTENT: Record<TaskpaneStatusType, MessageBarIntent> = {
  success: "success",
  error: "error",
};

const useStatusBarStyles = makeStyles({
  root: {
    flexShrink: 0,
    marginBottom: tokens.spacingVerticalS,
  },
});

/** Returns Griffel classes for the status bar. */
export function useStatusBar(): StatusBarClasses {
  const styles = useStatusBarStyles();
  return {
    root: styles.root,
  };
}

/** Maps the taskpane status type to the matching Fluent MessageBar intent. */
export function getStatusBarIntent(type: TaskpaneStatusType): MessageBarIntent {
  return STATUS_INTENT[type] ?? "info";
}
