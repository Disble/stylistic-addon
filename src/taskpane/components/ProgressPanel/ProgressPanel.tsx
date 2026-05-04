import * as React from "react";
import type { ProgressPanelProps } from "./ProgressPanel.types";

/** Renders the progress area used by the legacy pipeline observer. */
export function ProgressPanel({ progress }: ProgressPanelProps): React.JSX.Element {
  const progressPercent = progress.total > 0 ? Math.round((progress.current / progress.total) * 100) : 0;

  return (
    <div
      id="progress-container"
      className="progress-container"
      style={{ display: progress.visible ? "block" : "none" }}
    >
      <div className="progress-bar-track">
        <div id="progress-bar" className="progress-bar-fill" style={{ width: `${progressPercent}%` }} />
      </div>
      <p id="progress-text" className="progress-text">
        {progress.message}
      </p>
    </div>
  );
}
