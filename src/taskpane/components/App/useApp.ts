import { useTaskpaneShellState } from "../../TaskpaneShellStore";

/** Returns the current reactive shell state for the top-level App component. */
export function useApp() {
  return useTaskpaneShellState();
}
