import * as React from "react";
import { setTaskpaneSelectedGenero, useTaskpaneShellState } from "../../TaskpaneShellStore";

/** Returns the current reactive shell state for the top-level App component. */
export function useApp() {
  const shellState = useTaskpaneShellState();

  const handleGeneroChange = React.useCallback<React.ChangeEventHandler<HTMLSelectElement>>(
    (event) => {
      setTaskpaneSelectedGenero(event.target.value);
    },
    []
  );

  return {
    shellState,
    handleGeneroChange,
  };
}
