import { makeStyles, tokens } from "@fluentui/react-components";
import type { DisableTrackChangesSectionClasses } from "./DisableTrackChangesSection.types";

const useDisableTrackChangesSectionStyles = makeStyles({
  root: {
    flexShrink: 0,
    paddingBottom: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  button: {
    width: "100%",
    justifyContent: "center",
    color: tokens.colorPaletteDarkOrangeForeground1,
    borderTopColor: tokens.colorPaletteDarkOrangeBorder2,
    borderRightColor: tokens.colorPaletteDarkOrangeBorder2,
    borderBottomColor: tokens.colorPaletteDarkOrangeBorder2,
    borderLeftColor: tokens.colorPaletteDarkOrangeBorder2,
  },
});

/** Returns Griffel classes for the disable-track-changes section. */
export function useDisableTrackChangesSection(): DisableTrackChangesSectionClasses {
  const styles = useDisableTrackChangesSectionStyles();
  return {
    root: styles.root,
    button: styles.button,
  };
}
