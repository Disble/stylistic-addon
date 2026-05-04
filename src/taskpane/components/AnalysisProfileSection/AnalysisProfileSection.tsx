import * as React from "react";
import type { AnalysisProfileSectionProps } from "./AnalysisProfileSection.types";

/** Renders the analysis-profile selector while preserving the legacy DOM anchor IDs. */
export function AnalysisProfileSection({ options }: AnalysisProfileSectionProps): React.JSX.Element {
  return (
    <div className="stylistic-section">
      <label htmlFor="profile-select" className="stylistic-label">
        Perfil de análisis
      </label>
      <select id="profile-select" className="stylistic-select" defaultValue="narrativa-literaria">
        {options.map((option) => (
          <option key={option.value} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
    </div>
  );
}
