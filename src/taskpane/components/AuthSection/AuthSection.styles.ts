import { makeStyles, tokens } from "@fluentui/react-components";

/** Griffel styles for the authentication section. */
export const useAuthSectionStyles = makeStyles({
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
