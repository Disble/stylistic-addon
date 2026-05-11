import { useStatusBarStyles } from "./StatusBar.styles";
import type { StatusBarClasses } from "./StatusBar.types";

/** Returns Griffel classes for the status bar. */
export function useStatusBar(): StatusBarClasses {
  const styles = useStatusBarStyles();
  return {
    root: styles.root,
  };
}
