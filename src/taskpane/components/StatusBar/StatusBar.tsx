import * as React from "react";
import { MessageBar, MessageBarBody } from "@fluentui/react-components";
import type { StatusBarProps } from "./StatusBar.types";
import { getStatusBarIntent, useStatusBar } from "./useStatusBar";

/** Renders the bottom status MessageBar (success/error feedback for the user). */
export function StatusBar({ status }: StatusBarProps): React.JSX.Element | null {
  const classes = useStatusBar();
  if (!status.visible) {
    return null;
  }

  return (
    <MessageBar
      className={classes.root}
      data-testid="status-bar"
      intent={getStatusBarIntent(status.type)}
      politeness="polite"
    >
      <MessageBarBody>{status.message}</MessageBarBody>
    </MessageBar>
  );
}
