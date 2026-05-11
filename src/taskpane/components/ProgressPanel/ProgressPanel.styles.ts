import { makeStyles, tokens } from "@fluentui/react-components";

/** Griffel styles for the compact progress panel rendered above partial results. */
export const useProgressPanelStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalXS,
    marginBottom: tokens.spacingVerticalM,
    flexShrink: 0,
  },
  header: {
    display: "flex",
    alignItems: "center",
    columnGap: tokens.spacingHorizontalS,
  },
  message: {
    flex: "1 1 auto",
    color: tokens.colorNeutralForeground2,
    margin: 0,
    minWidth: 0,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  actions: {
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    columnGap: tokens.spacingHorizontalXS,
  },
});
