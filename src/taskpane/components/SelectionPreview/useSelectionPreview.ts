import { useResultsPanelStore } from "../../ResultsPanelStore";
import { useSelectionPreviewStore } from "../../SelectionPreviewStore";
import { useTaskpaneShellStore } from "../../TaskpaneShellStore";
import type { SelectionPreviewProps } from "./SelectionPreview.types";

/**
 * Derives the SelectionPreview view-state by combining selection, shell and
 * results stores. The preview is shown only when there is a real selection
 * AND the surface below the analyze button is otherwise empty.
 */
export function useSelectionPreview(): SelectionPreviewProps {
  const selection = useSelectionPreviewStore();
  const isAnalyzeLoading = useTaskpaneShellStore((state) => state.isAnalyzeLoading);
  const progressVisible = useTaskpaneShellStore((state) => state.progress.visible);
  const resultsVisible = useResultsPanelStore((state) => state.visible);

  const isVisible =
    selection.hasSelection && !isAnalyzeLoading && !progressVisible && !resultsVisible;

  return {
    isVisible,
    charCount: selection.charCount,
    preview: selection.preview,
  };
}
