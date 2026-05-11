import type { SelectionSnapshot } from "../domain/selection/SelectionSnapshot.types";

/** Initial empty selection snapshot used by the taskpane preview store. */
export const INITIAL_SELECTION_PREVIEW_STATE: SelectionSnapshot = {
  hasSelection: false,
  charCount: 0,
  preview: "",
};
