import { makeStyles, tokens } from "@fluentui/react-components";

/** Griffel styles for the account settings row. */
export const useAccountSettingsStyles = makeStyles({
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
