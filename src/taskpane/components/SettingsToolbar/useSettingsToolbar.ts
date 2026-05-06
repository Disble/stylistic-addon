import { makeStyles, tokens } from "@fluentui/react-components";
import type { SettingsToolbarClasses } from "./SettingsToolbar.types";

const useSettingsToolbarStyles = makeStyles({
  root: {
    display: "flex",
    justifyContent: "flex-end",
    alignItems: "center",
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalS,
    paddingRight: tokens.spacingHorizontalS,
    borderTopWidth: "1px",
    borderTopStyle: "solid",
    borderTopColor: tokens.colorNeutralStroke2,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  button: {
    minWidth: "auto",
  },
});

/** Returns Griffel classes for the bottom settings toolbar. */
export function useSettingsToolbar(): SettingsToolbarClasses {
  const styles = useSettingsToolbarStyles();
  return {
    root: styles.root,
    button: styles.button,
  };
}
