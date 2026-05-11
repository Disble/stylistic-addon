import { makeStyles, tokens } from "@fluentui/react-components";

/** Creates Griffel classes for the bottom settings toolbar. */
export const useSettingsToolbarStyles = makeStyles({
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
