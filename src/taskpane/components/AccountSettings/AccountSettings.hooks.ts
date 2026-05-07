import type { AccountSettingsClasses } from "./AccountSettings.types";
import { useAccountSettingsStyles } from "./AccountSettings.styles";

/** Returns Griffel classes for the account settings row. */
export function useAccountSettings(): AccountSettingsClasses {
  const styles = useAccountSettingsStyles();
  return {
    root: styles.root,
    header: styles.header,
    title: styles.title,
    row: styles.row,
    email: styles.email,
    logoutButton: styles.logoutButton,
  };
}
