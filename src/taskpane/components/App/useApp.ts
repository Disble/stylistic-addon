import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import { useTaskpaneAuthStore } from "../../TaskpaneAuthStore";
import { setTaskpaneSelectedGenero, useTaskpaneShellStore } from "../../TaskpaneShellStore";
import { setTaskpaneView, useTaskpaneViewStore } from "../../TaskpaneViewStore";
import type { AppClasses } from "./App.types";

const useAppStyles = makeStyles({
  workflow: {
    display: "flex",
    flexDirection: "column",
    flex: 1,
    minHeight: 0,
    overflowY: "auto",
  },
  toolbar: {
    flexShrink: 0,
    marginLeft: "-16px",
    marginRight: "-16px",
  },
});

/** Returns Griffel classes used by the App shell. */
export function useAppClasses(): AppClasses {
  const styles = useAppStyles();
  return {
    workflow: styles.workflow,
    toolbar: styles.toolbar,
  };
}

/**
 * Aggregates the reactive state and callbacks consumed by the App shell:
 * auth status, shell flags, active top-level view, and view-toggle handlers.
 * `handleOpenSettings` / `handleCloseSettings` flip `TaskpaneViewStore` so the
 * shell can re-render the appropriate surface (`main` vs `settings`).
 */
export function useApp() {
  const authState = useTaskpaneAuthStore();
  const shellState = useTaskpaneShellStore();
  const viewState = useTaskpaneViewStore();

  const handleGeneroChange = React.useCallback((value: string) => {
    setTaskpaneSelectedGenero(value);
  }, []);

  const handleOpenSettings = React.useCallback(() => {
    setTaskpaneView("settings");
  }, []);

  const handleCloseSettings = React.useCallback(() => {
    setTaskpaneView("main");
  }, []);

  return {
    authState,
    shellState,
    viewState,
    handleGeneroChange,
    handleOpenSettings,
    handleCloseSettings,
  };
}
