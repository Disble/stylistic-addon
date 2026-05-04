import * as React from "react";

/** Renders the progress area used by the legacy pipeline observer. */
export function ProgressPanel(): React.JSX.Element {
  return (
    <div id="progress-container" className="progress-container" style={{ display: "none" }}>
      <div className="progress-bar-track">
        <div id="progress-bar" className="progress-bar-fill" />
      </div>
      <p id="progress-text" className="progress-text" />
    </div>
  );
}
