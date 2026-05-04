import * as React from "react";

/** Renders the resolved-comments cleanup CTA. */
export function CleanupSection(): React.JSX.Element {
  return (
    <div className="stylistic-section" id="cleanup-section" style={{ display: "none" }}>
      <button id="btn-cleanup" type="button" className="stylistic-btn stylistic-btn--secondary">
        <span id="btn-cleanup-label">Limpiar comentarios resueltos</span>
      </button>
    </div>
  );
}
