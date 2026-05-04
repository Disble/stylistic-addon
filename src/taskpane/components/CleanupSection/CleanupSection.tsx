import * as React from "react";
import type { CleanupSectionProps } from "./CleanupSection.types";

/** Renders the resolved-comments cleanup CTA. */
export function CleanupSection({
  isLoading,
  isVisible,
  onCleanup,
}: CleanupSectionProps): React.JSX.Element {
  return (
    <div
      className="stylistic-section"
      id="cleanup-section"
      style={{ display: isVisible ? "block" : "none" }}
    >
      <button
        id="btn-cleanup"
        type="button"
        className="stylistic-btn stylistic-btn--secondary"
        disabled={isLoading}
        onClick={() => {
          void onCleanup();
        }}
      >
        <span id="btn-cleanup-label">
          {isLoading ? "Limpiando..." : "Limpiar comentarios resueltos"}
        </span>
      </button>
    </div>
  );
}
