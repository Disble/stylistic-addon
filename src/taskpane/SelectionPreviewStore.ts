import { create } from "zustand";
import type { SelectionSnapshot } from "../domain/selection/SelectionSnapshot.types";
import { INITIAL_SELECTION_PREVIEW_STATE } from "./SelectionPreviewStore.constants";

/** Zustand store holding the latest selection snapshot reported by the host. */
export const useSelectionPreviewStore = create<SelectionSnapshot>()(
  () => INITIAL_SELECTION_PREVIEW_STATE
);

/** Returns the current snapshot. */
export function getSelectionPreviewState(): SelectionSnapshot {
  return useSelectionPreviewStore.getState();
}

/** Replaces the snapshot with a new value emitted by the document port. */
export function setSelectionPreviewSnapshot(snapshot: SelectionSnapshot): void {
  useSelectionPreviewStore.setState(snapshot, true);
}

/** Resets the snapshot back to the empty state. */
export function resetSelectionPreviewState(): void {
  useSelectionPreviewStore.setState(INITIAL_SELECTION_PREVIEW_STATE, true);
}
