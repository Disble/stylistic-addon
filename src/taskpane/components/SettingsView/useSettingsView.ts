import { makeStyles, tokens } from "@fluentui/react-components";
import type { SettingsViewClasses } from "./SettingsView.types";

const useSettingsViewStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
  },
  header: {
    display: "grid",
    gridTemplateColumns: "auto 1fr auto",
    alignItems: "center",
    columnGap: tokens.spacingHorizontalS,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    borderBottomWidth: "1px",
    borderBottomStyle: "solid",
    borderBottomColor: tokens.colorNeutralStroke2,
  },
  backButton: {
    minWidth: "auto",
  },
  title: {
    textAlign: "center",
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  body: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalM,
    paddingTop: tokens.spacingVerticalM,
    flex: 1,
  },
});

/** Returns Griffel classes for the settings page. */
export function useSettingsView(): SettingsViewClasses {
  const styles = useSettingsViewStyles();
  return {
    root: styles.root,
    header: styles.header,
    backButton: styles.backButton,
    title: styles.title,
    body: styles.body,
  };
}
