import { makeStyles, tokens } from "@fluentui/react-components";
import { useSelectionPreviewStore } from "../../SelectionPreviewStore";
import { useTaskpaneShellStore } from "../../TaskpaneShellStore";
import type { SelectionPreviewProps } from "./SelectionPreview.types";

const useSelectionPreviewStyles = makeStyles({
  root: {
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    columnGap: tokens.spacingHorizontalS,
    paddingTop: tokens.spacingVerticalXS,
    paddingBottom: tokens.spacingVerticalXS,
    paddingLeft: tokens.spacingHorizontalS,
    paddingRight: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: "help",
  },
  icon: {
    flexShrink: 0,
    color: tokens.colorBrandForeground1,
    fontSize: tokens.fontSizeBase400,
    lineHeight: tokens.lineHeightBase400,
  },
  label: {
    flex: 1,
    minWidth: 0,
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
});

/**
 * Derives the SelectionPreview view-state by combining selection and shell stores
 * and resolves the Fluent UI styles. The preview is shown whenever there is a
 * real selection AND the analyze flow is not currently running. Results
 * visibility is intentionally NOT a gate: the user can re-analyze a new
 * fragment while previous suggestions remain on screen.
 */
export function useSelectionPreview(): SelectionPreviewProps {
  const selection = useSelectionPreviewStore();
  const isAnalyzeLoading = useTaskpaneShellStore((state) => state.isAnalyzeLoading);
  const progressVisible = useTaskpaneShellStore((state) => state.progress.visible);
  const classes = useSelectionPreviewStyles();

  const isVisible = selection.hasSelection && !isAnalyzeLoading && !progressVisible;

  return {
    isVisible,
    charCount: selection.charCount,
    preview: selection.preview,
    classes,
  };
}
