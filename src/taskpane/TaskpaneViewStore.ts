import { create } from "zustand";
import { INITIAL_TASKPANE_VIEW_STATE } from "./TaskpaneViewStore.constants";
import type { TaskpaneView, TaskpaneViewState } from "./TaskpaneViewStore.types";

/** Zustand store for the active top-level taskpane view. */
export const useTaskpaneViewStore = create<TaskpaneViewState>()(() => INITIAL_TASKPANE_VIEW_STATE);

/** Returns the current immutable view-state snapshot. */
export function getTaskpaneViewState(): TaskpaneViewState {
  return useTaskpaneViewStore.getState();
}

/** Switches the active top-level view. */
export function setTaskpaneView(view: TaskpaneView): void {
  useTaskpaneViewStore.setState({ view });
}

/** Resets view state for deterministic tests. */
export function resetTaskpaneViewState(): void {
  useTaskpaneViewStore.setState(INITIAL_TASKPANE_VIEW_STATE, true);
}
