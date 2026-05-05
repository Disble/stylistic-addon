import { makeStyles, tokens } from "@fluentui/react-components";
import type { ProgressPanelClasses } from "./ProgressPanel.types";

const useProgressPanelStyles = makeStyles({
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

/** Returns Griffel classes for the progress panel. */
export function useProgressPanel(): ProgressPanelClasses {
  const styles = useProgressPanelStyles();
  return {
    root: styles.root,
    message: styles.message,
  };
}
