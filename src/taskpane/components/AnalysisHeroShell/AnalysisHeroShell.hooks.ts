import type { AnalysisHeroShellClasses } from "./AnalysisHeroShell.types";
import { useAnalysisHeroShellStyles } from "./AnalysisHeroShell.styles";

/** Returns Griffel classes consumed by the shared analysis hero shell. */
export function useAnalysisHeroShell(): AnalysisHeroShellClasses {
  const styles = useAnalysisHeroShellStyles();
  return {
    root: styles.root,
    illustrationWrapper: styles.illustrationWrapper,
    copy: styles.copy,
    title: styles.title,
    message: styles.message,
    actions: styles.actions,
  };
}
