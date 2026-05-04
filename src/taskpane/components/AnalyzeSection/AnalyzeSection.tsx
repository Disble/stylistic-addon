import * as React from "react";

/** Renders the primary analysis CTA. */
export function AnalyzeSection(): React.JSX.Element {
  return (
    <div className="stylistic-section">
      <button id="btn-analyze" type="button" className="stylistic-btn stylistic-btn--primary">
        <span id="btn-analyze-label">Analizar y sugerir</span>
      </button>
    </div>
  );
}
