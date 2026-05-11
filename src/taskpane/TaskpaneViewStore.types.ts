/** Top-level taskpane route represented in React-owned presentation state. */
export type TaskpaneView = "main" | "settings";

/** Root Zustand state for taskpane view selection. */
export type TaskpaneViewState = Readonly<{
  view: TaskpaneView;
}>;
