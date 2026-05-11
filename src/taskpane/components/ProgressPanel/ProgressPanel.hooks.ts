import type { ProgressPanelClasses } from "./ProgressPanel.types";
import { useProgressPanelStyles } from "./ProgressPanel.styles";

/** Returns Griffel classes for the compact progress panel. */
export function useProgressPanel(): ProgressPanelClasses {
  const styles = useProgressPanelStyles();
  return {
    root: styles.root,
    header: styles.header,
    message: styles.message,
    actions: styles.actions,
  };
}
