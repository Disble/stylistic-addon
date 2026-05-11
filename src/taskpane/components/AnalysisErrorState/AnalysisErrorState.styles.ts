import { makeStyles, tokens } from "@fluentui/react-components";

const heroIn = {
  from: { opacity: 0, transform: "translateY(8px)" },
  to: { opacity: 1, transform: "translateY(0)" },
};

const alertPulse = {
  "0%, 100%": { transform: "scale(1)" },
  "50%": { transform: "scale(1.06)" },
};

/** Creates Griffel classes for the hero-style analysis-error surface and illustration. */
export const useAnalysisErrorStateStyles = makeStyles({
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
  illustration: {
    width: "200px",
    maxWidth: "100%",
    height: "auto",
    flexShrink: 0,
  },
  illustrationDoc: {
    color: tokens.colorBrandForeground1,
  },
  illustrationAlert: {
    color: tokens.colorPaletteRedBackground3,
    transformOrigin: "center",
    transformBox: "fill-box",
    animationName: alertPulse,
    animationDuration: "2400ms",
    animationIterationCount: "infinite",
    animationTimingFunction: "ease-in-out",
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
  guidance: {
    margin: 0,
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
  actions: {
    flexShrink: 0,
    display: "flex",
    flexDirection: "column",
    alignItems: "stretch",
    rowGap: tokens.spacingVerticalS,
  },
  retryButton: {
    boxSizing: "border-box",
  },
});
