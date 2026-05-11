import type { AnalysisErrorStateClasses } from "./AnalysisErrorState.types";
import { useAnalysisErrorStateStyles } from "./AnalysisErrorState.styles";

/** Returns Griffel classes consumed by the hero-style analysis-error surface. */
export function useAnalysisErrorState(): AnalysisErrorStateClasses {
  const styles = useAnalysisErrorStateStyles();
  return {
    root: styles.root,
    illustrationWrapper: styles.illustrationWrapper,
    illustration: styles.illustration,
    illustrationDoc: styles.illustrationDoc,
    illustrationAlert: styles.illustrationAlert,
    copy: styles.copy,
    title: styles.title,
    message: styles.message,
    guidance: styles.guidance,
    actions: styles.actions,
    retryButton: styles.retryButton,
  };
}
