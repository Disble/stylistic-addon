import type { SelectionSnapshot } from "../../domain/selection/SelectionSnapshot.types";

/** Default selection preview length exposed by the monitor adapter. */
export const DEFAULT_PREVIEW_MAX_CHARS = 80;

/** Empty snapshot emitted when Word has no active text selection. */
export const EMPTY_SELECTION_SNAPSHOT: SelectionSnapshot = {
  hasSelection: false,
  charCount: 0,
  preview: "",
};
