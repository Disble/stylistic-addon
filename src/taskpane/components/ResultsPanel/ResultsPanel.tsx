import * as React from "react";

/** Renders the results panel anchors consumed by the legacy DOM renderer. */
export function ResultsPanel(): React.JSX.Element {
  return (
    <div id="results-panel" className="results-panel" style={{ display: "none" }}>
      <h2 className="results-title">Resultados</h2>
      <div id="results-summary" className="results-summary" />
      <ul id="results-list" className="results-list" />
    </div>
  );
}
