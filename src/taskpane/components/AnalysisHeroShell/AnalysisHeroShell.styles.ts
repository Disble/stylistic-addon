import { makeStyles, tokens } from "@fluentui/react-components";

const heroIn = {
  from: { opacity: 0, transform: "translateY(8px)" },
  to: { opacity: 1, transform: "translateY(0)" },
};

/** Creates Griffel classes for the shared analysis hero shell layout. */
export const useAnalysisHeroShellStyles = makeStyles({
  root: {
    flex: "1 1 auto",
    display: "flex",
    flexDirection: "column",
    minHeight: 0,
    paddingTop: tokens.spacingVerticalL,
    paddingBottom: tokens.spacingVerticalL,
    rowGap: tokens.spacingVerticalM,
    animationName: heroIn,
    animationDuration: tokens.durationGentle,
    animationTimingFunction: tokens.curveDecelerateMid,
    animationFillMode: "both",
  },
  illustrationWrapper: {
    flex: "1 1 auto",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalM,
    minHeight: 0,
  },
  copy: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    rowGap: tokens.spacingVerticalXS,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
  },
  title: {
    margin: 0,
    textAlign: "center",
    color: tokens.colorNeutralForeground1,
  },
  message: {
    margin: 0,
    textAlign: "center",
    color: tokens.colorNeutralForeground2,
  },
  actions: {
    flexShrink: 0,
    display: "flex",
    flexDirection: "column",
    alignItems: "stretch",
    rowGap: tokens.spacingVerticalS,
  },
});
