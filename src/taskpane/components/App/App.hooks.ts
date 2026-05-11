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
 * Computes mutually-exclusive surface flags (`isIdle`, `isHeroError`,
 * `isHeroProgress`) so the App template chooses one fullscreen surface at a
 * time and avoids stacking the compact analyze CTA on top of a hero state.
 */
export function useApp() {
  const authState = useTaskpaneAuthStore();
  const shellState = useTaskpaneShellStore();
  const viewState = useTaskpaneViewStore();
  const resultsVisible = useResultsPanelStore((state) => state.visible);
  const hasResultCards = useResultsPanelStore((state) => state.cards.length > 0);

  const isHeroError =
    shellState.analysisError.visible &&
    !shellState.isAnalyzeLoading &&
    !shellState.progress.visible &&
    !resultsVisible;

  const isHeroProgress =
    !shellState.analysisError.visible &&
    (shellState.progress.visible || shellState.isAnalyzeLoading) &&
    !hasResultCards;

  const isIdle =
    !shellState.isAnalyzeLoading &&
    !shellState.progress.visible &&
    !shellState.analysisError.visible &&
    !resultsVisible;

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
    isHeroError,
    isHeroProgress,
    handleOpenSettings,
    handleCloseSettings,
  };
}
