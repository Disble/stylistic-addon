import * as React from "react";
import type { AnalyzeSectionProps } from "./AnalyzeSection.types";

/** Renders the primary analysis CTA. */
export function AnalyzeSection({ isLoading }: AnalyzeSectionProps): React.JSX.Element {
  return (
    <div className="stylistic-section">
      <button
        id="btn-analyze"
        type="button"
        className="stylistic-btn stylistic-btn--primary"
        disabled={isLoading}
      >
        <span id="btn-analyze-label">{isLoading ? "Analizando..." : "Analizar y sugerir"}</span>
      </button>
    </div>
  );
}
