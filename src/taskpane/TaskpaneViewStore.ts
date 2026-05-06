import { create } from "zustand";

export type TaskpaneView = "main" | "settings";

export type TaskpaneViewState = Readonly<{
  view: TaskpaneView;
}>;

const INITIAL_STATE: TaskpaneViewState = {
  view: "main",
};

/** Zustand store for the active top-level taskpane view. */
export const useTaskpaneViewStore = create<TaskpaneViewState>()(() => INITIAL_STATE);

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
  useTaskpaneViewStore.setState(INITIAL_STATE, true);
}
