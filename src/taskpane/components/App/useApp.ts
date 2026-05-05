import * as React from "react";
import { setTaskpaneSelectedGenero, useTaskpaneShellStore } from "../../TaskpaneShellStore";

/** Returns the current reactive shell state for the top-level App component. */
export function useApp() {
  const shellState = useTaskpaneShellStore();

  const handleGeneroChange = React.useCallback((value: string) => {
    setTaskpaneSelectedGenero(value);
  }, []);

  return {
    shellState,
    handleGeneroChange,
  };
}
