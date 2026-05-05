import { makeStyles, tokens } from "@fluentui/react-components";
import type { AnalysisProfileSectionClasses } from "./AnalysisProfileSection.types";

const useAnalysisProfileSectionStyles = makeStyles({
  root: {
    marginBottom: tokens.spacingVerticalM,
  },
  field: {
    width: "100%",
  },
  dropdown: {
    width: "100%",
  },
});

/** Returns Griffel classes for the analysis-profile section. */
export function useAnalysisProfileSection(): AnalysisProfileSectionClasses {
  const styles = useAnalysisProfileSectionStyles();
  return {
    root: styles.root,
    field: styles.field,
    dropdown: styles.dropdown,
  };
}
