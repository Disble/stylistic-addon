import * as React from "react";
import { useTaskpaneAuthStore } from "../../TaskpaneAuthStore";
import { setTaskpaneSelectedGenero, useTaskpaneShellStore } from "../../TaskpaneShellStore";

/** Returns the current reactive shell state for the top-level App component. */
export function useApp() {
  const authState = useTaskpaneAuthStore();
  const shellState = useTaskpaneShellStore();

  const handleGeneroChange = React.useCallback((value: string) => {
    setTaskpaneSelectedGenero(value);
  }, []);

  return {
    authState,
    shellState,
    handleGeneroChange,
  };
}
