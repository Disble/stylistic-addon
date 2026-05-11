import { makeStyles, tokens } from "@fluentui/react-components";

/** Griffel styles for the disable-track-changes section. */
export const useDisableTrackChangesSectionStyles = makeStyles({
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
