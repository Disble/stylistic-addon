import * as React from "react";
import { useResultsPanelStore } from "../../ResultsPanelStore";
import { useTaskpaneAuthStore } from "../../TaskpaneAuthStore";
import { useTaskpaneShellStore } from "../../TaskpaneShellStore";
import { setTaskpaneView, useTaskpaneViewStore } from "../../TaskpaneViewStore";
import { useAppStyles } from "./App.styles";
import type { AppClasses } from "./App.types";

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
  const resultsVisible = useResultsPanelStore((state) => state.visible);

  const isIdle = !shellState.isAnalyzeLoading && !shellState.progress.visible && !resultsVisible;

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
    isIdle,
    handleOpenSettings,
    handleCloseSettings,
  };
}
