import type { CleanupSectionClasses } from "./CleanupSection.types";
import { useCleanupSectionStyles } from "./CleanupSection.styles";

/** Returns Griffel classes for the cleanup section. */
export function useCleanupSection(): CleanupSectionClasses {
  const styles = useCleanupSectionStyles();
  return {
    root: styles.root,
    button: styles.button,
  };
}
