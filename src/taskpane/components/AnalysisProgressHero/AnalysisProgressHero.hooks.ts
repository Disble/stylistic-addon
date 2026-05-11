import type { AnalysisProgressHeroClasses } from "./AnalysisProgressHero.types";
import { useAnalysisProgressHeroStyles } from "./AnalysisProgressHero.styles";

/** Returns Griffel classes consumed by the hero-style analysis-progress surface. */
export function useAnalysisProgressHero(): AnalysisProgressHeroClasses {
  const styles = useAnalysisProgressHeroStyles();
  return {
    root: styles.root,
    illustrationWrapper: styles.illustrationWrapper,
    illustration: styles.illustration,
    illustrationDoc: styles.illustrationDoc,
    illustrationWand: styles.illustrationWand,
    illustrationSparkle: styles.illustrationSparkle,
    illustrationSparkle1: styles.illustrationSparkle1,
    illustrationSparkle2: styles.illustrationSparkle2,
    illustrationSparkle3: styles.illustrationSparkle3,
    illustrationLine: styles.illustrationLine,
    illustrationLineDelay1: styles.illustrationLineDelay1,
    illustrationLineDelay2: styles.illustrationLineDelay2,
    illustrationLineDelay3: styles.illustrationLineDelay3,
    illustrationLineDelay4: styles.illustrationLineDelay4,
    illustrationLineDelay5: styles.illustrationLineDelay5,
    illustrationLineDelay6: styles.illustrationLineDelay6,
    copy: styles.copy,
    title: styles.title,
    message: styles.message,
    progressField: styles.progressField,
    actions: styles.actions,
    primaryButton: styles.primaryButton,
  };
}
