import { useSettingsSaveBarStyles } from "./SettingsSaveBar.styles";
import type { SettingsSaveBarClasses } from "./SettingsSaveBar.types";

/** Returns Griffel classes for the settings save bar. */
export function useSettingsSaveBar(): SettingsSaveBarClasses {
  const styles = useSettingsSaveBarStyles();
  return {
    root: styles.root,
    error: styles.error,
    actions: styles.actions,
    saveButton: styles.saveButton,
  };
}
