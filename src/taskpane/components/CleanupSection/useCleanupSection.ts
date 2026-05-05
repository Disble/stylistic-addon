import { makeStyles, tokens } from "@fluentui/react-components";
import type { CleanupSectionClasses } from "./CleanupSection.types";

const useCleanupSectionStyles = makeStyles({
  root: {
    flexShrink: 0,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalM,
    borderTopWidth: "1px",
    borderTopStyle: "solid",
    borderTopColor: tokens.colorNeutralStroke2,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  button: {
    width: "100%",
    justifyContent: "center",
  },
});

/** Returns Griffel classes for the cleanup section. */
export function useCleanupSection(): CleanupSectionClasses {
  const styles = useCleanupSectionStyles();
  return {
    root: styles.root,
    button: styles.button,
  };
}
