import * as React from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import { DEFAULT_PROFILES } from "../../../infrastructure/config";
import { setTaskpaneSelectedGenero, useTaskpaneShellStore } from "../../TaskpaneShellStore";
import type { AnalysisProfileOption } from "../AnalysisProfileSection";
import type { SettingsViewClasses } from "./SettingsView.types";

const ANALYSIS_PROFILE_OPTIONS: readonly AnalysisProfileOption[] = DEFAULT_PROFILES.map(
  (profile) => ({
    value: profile.id,
    label: profile.label,
  })
);

const useSettingsViewStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
  },
  header: {
    display: "grid",
    gridTemplateColumns: "auto 1fr auto",
    alignItems: "center",
    columnGap: tokens.spacingHorizontalS,
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    borderBottomWidth: "1px",
    borderBottomStyle: "solid",
    borderBottomColor: tokens.colorNeutralStroke2,
  },
  backButton: {
    minWidth: "auto",
  },
  title: {
    textAlign: "center",
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  body: {
    display: "flex",
    flexDirection: "column",
    rowGap: tokens.spacingVerticalM,
    paddingTop: tokens.spacingVerticalM,
    flex: 1,
  },
});

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

export type SettingsViewState = Readonly<{
  classes: SettingsViewClasses;
  analysisProfileOptions: readonly AnalysisProfileOption[];
  isAnalysisProfileDisabled: boolean;
  selectedGenero: string;
  handleGeneroChange: (value: string) => void;
}>;

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
