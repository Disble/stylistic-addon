import { makeStyles, tokens } from "@fluentui/react-components";

/** Griffel styles for the progress panel. */
export const useProgressPanelStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalXS,
    marginBottom: tokens.spacingVerticalM,
    flexShrink: 0,
  },
  message: {
    color: tokens.colorNeutralForeground2,
    textAlign: "center",
    margin: 0,
  },
});
