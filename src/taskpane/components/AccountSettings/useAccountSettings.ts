import { makeStyles, tokens } from "@fluentui/react-components";
import type { AccountSettingsClasses } from "./AccountSettings.types";

const useAccountSettingsStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalS,
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
  },
  header: {
    display: "flex",
    alignItems: "center",
    columnGap: tokens.spacingHorizontalXS,
    color: tokens.colorNeutralForeground2,
  },
  title: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  row: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    columnGap: tokens.spacingHorizontalM,
  },
  email: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground1,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  logoutButton: {
    flexShrink: 0,
  },
});

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
