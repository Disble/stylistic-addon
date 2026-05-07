import * as React from "react";
import { setTaskpaneSelectedGenero, useTaskpaneShellStore } from "../../TaskpaneShellStore";
import { ANALYSIS_PROFILE_OPTIONS } from "./SettingsView.constants";
import { useSettingsViewStyles } from "./SettingsView.styles";
import type { SettingsViewClasses, SettingsViewState } from "./SettingsView.types";

/** Returns Griffel classes for the settings page. */
export function useSettingsViewClasses(): SettingsViewClasses {
  const styles = useSettingsViewStyles();
  return {
    root: styles.root,
    header: styles.header,
    backButton: styles.backButton,
    title: styles.title,
    body: styles.body,
  };
}

/**
 * Aggregates Settings page state: classes, the analysis-profile options
 * derived from the canonical domain profile list, the currently selected
 * profile, and the change handler that persists the selection back to the
 * shell store. The selector is disabled while an analysis run is active so the
 * persisted preference cannot drift away from the pipeline snapshot already in
 * flight.
 */
export function useSettingsView(): SettingsViewState {
  const classes = useSettingsViewClasses();
  const selectedGenero = useTaskpaneShellStore((state) => state.selectedGenero);
  const isAnalysisProfileDisabled = useTaskpaneShellStore((state) => state.isAnalyzeLoading);

  const handleGeneroChange = React.useCallback((value: string) => {
    setTaskpaneSelectedGenero(value);
  }, []);

  return {
    classes,
    analysisProfileOptions: ANALYSIS_PROFILE_OPTIONS,
    isAnalysisProfileDisabled,
    selectedGenero,
    handleGeneroChange,
  };
}
