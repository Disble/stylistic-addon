import { makeStyles } from "@fluentui/react-components";

/** Creates Griffel classes for the top-level taskpane shell. */
export const useAppStyles = makeStyles({
  workflow: {
    display: "flex",
    flexDirection: "column",
    flex: 1,
    minHeight: 0,
    overflowX: "hidden",
    overflowY: "auto",
  },
  toolbar: {
    flexShrink: 0,
    marginLeft: "-16px",
    marginRight: "-16px",
  },
});
