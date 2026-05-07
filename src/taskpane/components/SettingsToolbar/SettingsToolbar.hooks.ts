import type { SettingsToolbarClasses } from "./SettingsToolbar.types";
import { useSettingsToolbarStyles } from "./SettingsToolbar.styles";

/** Returns Griffel classes for the bottom settings toolbar. */
export function useSettingsToolbar(): SettingsToolbarClasses {
  const styles = useSettingsToolbarStyles();
  return {
    root: styles.root,
    button: styles.button,
  };
}
