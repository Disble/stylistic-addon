import * as React from "react";
import type { AnalyzeSectionProps } from "./AnalyzeSection.types";

/** Renders the primary analysis CTA. */
export function AnalyzeSection({ isLoading, onAnalyze }: AnalyzeSectionProps): React.JSX.Element {
  return (
    <div className="stylistic-section">
      <button
        id="btn-analyze"
        type="button"
        className="stylistic-btn stylistic-btn--primary"
        disabled={isLoading}
        onClick={() => {
          void onAnalyze();
        }}
      >
        <span id="btn-analyze-label">{isLoading ? "Analizando..." : "Analizar y sugerir"}</span>
      </button>
    </div>
  );
}
