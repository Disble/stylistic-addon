import * as React from "react";
import type { StatusBarProps } from "./StatusBar.types";

/** Renders the bottom status-bar anchor. */
export function StatusBar({ status }: StatusBarProps): React.JSX.Element {
  return (
    <div
      id="status-bar"
      className={`stylistic-status ${status.type}`}
      style={{ display: status.visible ? "block" : "none" }}
    >
      {status.message}
    </div>
  );
}
