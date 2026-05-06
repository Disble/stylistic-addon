import { makeStyles, tokens } from "@fluentui/react-components";
import type { HeroEmptyStateClasses } from "./HeroEmptyState.types";

const heroIn = {
  from: { opacity: 0, transform: "translateY(8px)" },
  to: { opacity: 1, transform: "translateY(0)" },
};

const sparklePulse = {
  "0%, 100%": { opacity: 0.35, transform: "scale(0.85)" },
  "50%": { opacity: 1, transform: "scale(1.1)" },
};

const useHeroEmptyStateStyles = makeStyles({
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
    color: tokens.colorBrandForeground1,
    flexShrink: 0,
  },
  sparkle1: {
    transformOrigin: "center",
    transformBox: "fill-box",
    animationName: sparklePulse,
    animationDuration: "2200ms",
    animationIterationCount: "infinite",
    animationTimingFunction: "ease-in-out",
  },
  sparkle2: {
    transformOrigin: "center",
    transformBox: "fill-box",
    animationName: sparklePulse,
    animationDuration: "2600ms",
    animationDelay: "400ms",
    animationIterationCount: "infinite",
    animationTimingFunction: "ease-in-out",
  },
  sparkle3: {
    transformOrigin: "center",
    transformBox: "fill-box",
    animationName: sparklePulse,
    animationDuration: "1800ms",
    animationDelay: "900ms",
    animationIterationCount: "infinite",
    animationTimingFunction: "ease-in-out",
  },
  title: {
    margin: 0,
    textAlign: "center",
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
    lineHeight: tokens.lineHeightBase400,
  },
  subtitle: {
    margin: 0,
    textAlign: "center",
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
  },
  actions: {
    flexShrink: 0,
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalS,
  },
});

/** Returns Griffel classes for the hero empty-state surface. */
export function useHeroEmptyState(): HeroEmptyStateClasses {
  const styles = useHeroEmptyStateStyles();
  return {
    root: styles.root,
    illustrationWrapper: styles.illustrationWrapper,
    illustration: styles.illustration,
    sparkle1: styles.sparkle1,
    sparkle2: styles.sparkle2,
    sparkle3: styles.sparkle3,
    title: styles.title,
    subtitle: styles.subtitle,
    actions: styles.actions,
  };
}
