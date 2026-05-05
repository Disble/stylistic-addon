import { makeStyles, tokens } from "@fluentui/react-components";
import type { AnalyzeSectionClasses } from "./AnalyzeSection.types";

const useAnalyzeSectionStyles = makeStyles({
  root: {
    marginBottom: tokens.spacingVerticalM,
  },
  button: {
    width: "100%",
    justifyContent: "center",
  },
});

/** Returns Griffel classes for the analyze CTA section. */
export function useAnalyzeSection(): AnalyzeSectionClasses {
  const styles = useAnalyzeSectionStyles();
  return {
    root: styles.root,
    button: styles.button,
  };
}
