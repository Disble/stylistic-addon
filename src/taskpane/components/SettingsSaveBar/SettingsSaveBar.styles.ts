import { makeStyles, tokens } from "@fluentui/react-components";

/** Griffel styles for the settings save bar. */
export const useSettingsSaveBarStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalXS,
    paddingTop: tokens.spacingVerticalM,
    borderTopWidth: "1px",
    borderTopStyle: "solid",
    borderTopColor: tokens.colorNeutralStroke2,
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  actions: {
    display: "flex",
    justifyContent: "flex-end",
  },
  saveButton: {
    minWidth: "120px",
  },
});
