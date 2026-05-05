import { makeStyles, tokens } from "@fluentui/react-components";
import type { AuthSectionClasses } from "./AuthSection.types";

const useAuthSectionStyles = makeStyles({
  root: {
    paddingTop: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalM,
    borderBottomWidth: "1px",
    borderBottomStyle: "solid",
    borderBottomColor: tokens.colorNeutralStroke2,
  },
  content: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalS,
  },
  title: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  description: {
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase200,
  },
  button: {
    width: "100%",
    justifyContent: "center",
  },
});

/** Returns Griffel classes for the authentication section. */
export function useAuthSection(): AuthSectionClasses {
  const styles = useAuthSectionStyles();
  return {
    root: styles.root,
    content: styles.content,
    title: styles.title,
    description: styles.description,
    button: styles.button,
  };
}
