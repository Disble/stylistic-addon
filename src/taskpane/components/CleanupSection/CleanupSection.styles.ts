import { makeStyles, tokens } from "@fluentui/react-components";

/** Griffel styles for the cleanup section. */
export const useCleanupSectionStyles = makeStyles({
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
