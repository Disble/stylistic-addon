import type { ProgressPanelClasses } from "./ProgressPanel.types";
import { useProgressPanelStyles } from "./ProgressPanel.styles";

/** Returns Griffel classes for the progress panel. */
export function useProgressPanel(): ProgressPanelClasses {
  const styles = useProgressPanelStyles();
  return {
    root: styles.root,
    message: styles.message,
  };
}
