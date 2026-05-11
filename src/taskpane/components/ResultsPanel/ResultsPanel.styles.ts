import { makeStyles, tokens } from "@fluentui/react-components";

/** Creates Griffel classes for the results panel list and empty state. */
export const useResultsPanelStyles = makeStyles({
  root: {
    flex: "1 1 auto",
    minHeight: 0,
    overflowY: "auto",
    paddingBottom: tokens.spacingVerticalS,
  },
  list: {
    listStyle: "none",
    margin: 0,
    paddingTop: tokens.spacingVerticalS,
    paddingRight: 0,
    paddingBottom: 0,
    paddingLeft: 0,
  },
  empty: {
    paddingTop: tokens.spacingVerticalL,
    paddingBottom: tokens.spacingVerticalL,
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
});
