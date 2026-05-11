import { makeStyles, tokens } from "@fluentui/react-components";

const heroIn = {
  from: { opacity: 0, transform: "translateY(8px)" },
  to: { opacity: 1, transform: "translateY(0)" },
};

const wandWiggle = {
  "0%, 100%": { transform: "rotate(-6deg)" },
  "50%": { transform: "rotate(6deg)" },
};

const sparklePulseFast = {
  "0%, 100%": { opacity: 0.4, transform: "scale(0.85)" },
  "50%": { opacity: 1, transform: "scale(1.2)" },
};

const linePulse = {
  "0%, 100%": { opacity: 0.18 },
  "50%": { opacity: 0.42 },
};

/** Creates Griffel classes for the hero-style analysis-progress surface and illustration. */
export const useAnalysisProgressHeroStyles = makeStyles({
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
  illustrationWand: {
    transformOrigin: "190px 22px",
    transformBox: "fill-box",
    animationName: wandWiggle,
    animationDuration: "1800ms",
    animationIterationCount: "infinite",
    animationTimingFunction: "ease-in-out",
  },
  illustrationSparkle: {
    transformOrigin: "center",
    transformBox: "fill-box",
    animationName: sparklePulseFast,
    animationIterationCount: "infinite",
    animationTimingFunction: "ease-in-out",
  },
  illustrationSparkle1: {
    animationDuration: "1400ms",
  },
  illustrationSparkle2: {
    animationDuration: "1600ms",
    animationDelay: "300ms",
  },
  illustrationSparkle3: {
    animationDuration: "1200ms",
    animationDelay: "600ms",
  },
  illustrationLine: {
    animationName: linePulse,
    animationIterationCount: "infinite",
    animationTimingFunction: "ease-in-out",
    animationDuration: "1800ms",
  },
  illustrationLineDelay1: { animationDelay: "0ms" },
  illustrationLineDelay2: { animationDelay: "120ms" },
  illustrationLineDelay3: { animationDelay: "240ms" },
  illustrationLineDelay4: { animationDelay: "360ms" },
  illustrationLineDelay5: { animationDelay: "480ms" },
  illustrationLineDelay6: { animationDelay: "600ms" },
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
    color: tokens.colorNeutralForeground3,
  },
  progressField: {
    alignSelf: "stretch",
    boxSizing: "border-box",
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
  },
  actions: {
    flexShrink: 0,
    display: "flex",
    flexDirection: "column",
    alignItems: "stretch",
    rowGap: tokens.spacingVerticalS,
  },
  primaryButton: {
    boxSizing: "border-box",
  },
});
