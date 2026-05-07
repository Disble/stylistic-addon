import type { AnalysisProfileSectionClasses } from "./AnalysisProfileSection.types";
import { useAnalysisProfileSectionStyles } from "./AnalysisProfileSection.styles";

/** Returns Griffel classes for the analysis-profile section. */
export function useAnalysisProfileSection(): AnalysisProfileSectionClasses {
  const styles = useAnalysisProfileSectionStyles();
  return {
    root: styles.root,
    field: styles.field,
    dropdown: styles.dropdown,
  };
}
