import { create } from "zustand";
import type { SelectionSnapshot } from "../domain/selection/SelectionSnapshot.types";

const INITIAL_STATE: SelectionSnapshot = {
  hasSelection: false,
  charCount: 0,
  preview: "",
};

/** Zustand store holding the latest selection snapshot reported by the host. */
export const useSelectionPreviewStore = create<SelectionSnapshot>()(() => INITIAL_STATE);

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
  useSelectionPreviewStore.setState(INITIAL_STATE, true);
}
