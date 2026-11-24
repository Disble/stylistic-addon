import type { AnalyzeSectionClasses } from "./AnalyzeSection.types";
import { useAnalyzeSectionStyles } from "./AnalyzeSection.styles";

/** Returns Griffel classes for the analyze CTA section. */
export function useAnalyzeSection(): AnalyzeSectionClasses {
  const styles = useAnalyzeSectionStyles();
  return {
    root: styles.root,
    button: styles.button,
  };
}
