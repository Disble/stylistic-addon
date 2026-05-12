import { makeStyles, tokens } from "@fluentui/react-components";

const alertPulse = {
  "0%, 100%": { transform: "scale(1)" },
  "50%": { transform: "scale(1.06)" },
};

/** Creates Griffel classes for the hero-style analysis-error surface and illustration. */
export const useAnalysisErrorStateStyles = makeStyles({
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
  guidance: {
    margin: 0,
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
  },
  retryButton: {
    boxSizing: "border-box",
  },
});
