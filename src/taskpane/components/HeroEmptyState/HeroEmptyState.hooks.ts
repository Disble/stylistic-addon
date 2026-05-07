import { useHeroEmptyStateStyles } from "./HeroEmptyState.styles";
import type { HeroEmptyStateClasses } from "./HeroEmptyState.types";

/** Returns Griffel classes for the hero empty-state surface. */
export function useHeroEmptyState(): HeroEmptyStateClasses {
  const styles = useHeroEmptyStateStyles();
  return {
    root: styles.root,
    illustrationWrapper: styles.illustrationWrapper,
    illustration: styles.illustration,
    sparkle1: styles.sparkle1,
    sparkle2: styles.sparkle2,
    sparkle3: styles.sparkle3,
    title: styles.title,
    subtitle: styles.subtitle,
    actions: styles.actions,
  };
}
