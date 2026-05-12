import type { AnalysisErrorStateClasses } from "./AnalysisErrorState.types";
import { useAnalysisErrorStateStyles } from "./AnalysisErrorState.styles";

/** Returns Griffel classes consumed by the hero-style analysis-error surface. */
export function useAnalysisErrorState(): AnalysisErrorStateClasses {
  const styles = useAnalysisErrorStateStyles();
  return {
    illustration: styles.illustration,
    illustrationDoc: styles.illustrationDoc,
    illustrationAlert: styles.illustrationAlert,
    guidance: styles.guidance,
    retryButton: styles.retryButton,
  };
}
