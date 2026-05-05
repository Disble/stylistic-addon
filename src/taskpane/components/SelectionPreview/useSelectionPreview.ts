import { makeStyles, tokens } from "@fluentui/react-components";
import { useSelectionPreviewStore } from "../../SelectionPreviewStore";
import { useTaskpaneShellStore } from "../../TaskpaneShellStore";
import type { SelectionPreviewProps } from "./SelectionPreview.types";

const useSelectionPreviewStyles = makeStyles({
  root: {
    marginBottom: tokens.spacingVerticalM,
  },
  body: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalXS,
  },
  quote: {
    fontStyle: "italic",
    color: tokens.colorNeutralForeground2,
    overflowWrap: "anywhere",
  },
  hint: {
    color: tokens.colorBrandForeground1,
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
