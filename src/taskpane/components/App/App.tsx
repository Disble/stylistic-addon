import * as React from "react";
import { AnalysisErrorState } from "../AnalysisErrorState";
import { AnalysisProgressHero } from "../AnalysisProgressHero";
import { AnalyzeSection } from "../AnalyzeSection";
import { CleanupSection } from "../CleanupSection";
import { DisableTrackChangesSection } from "../DisableTrackChangesSection";
import { HeroEmptyState } from "../HeroEmptyState";
import { ProgressPanel } from "../ProgressPanel";
import { ResultsPanel } from "../ResultsPanel";
import { SelectionPreview, useSelectionPreview } from "../SelectionPreview";
import { SettingsToolbar } from "../SettingsToolbar";
import { SettingsView } from "../SettingsView";
import { StatusBar } from "../StatusBar";
import type { AppProps } from "./App.types";
import { useApp, useAppClasses } from "./App.hooks";

/**
 * React shell for the taskpane.
 *
 * Picks one of four workflow surfaces based on shell state:
 * - `isIdle` → `HeroEmptyState` with the analyze CTA
 * - `isHeroProgress` → `AnalysisProgressHero` while analyzing and no cards exist
 * - `isHeroError` → `AnalysisErrorState` when an analysis attempt failed
 * - otherwise → analyze CTA + compact progress strip + results list
 *
 * Preserves the legacy `#app-body` DOM anchor so `taskpane.css` can keep owning the
 * outer flex layout (height chain + padding) without a port through Griffel.
 */
export function App({
  onAnalyze,
  onRetryAnalysis,
  onCancelAnalysis,
  onRetryAnalysisQuery,
  onCleanup,
  onDisableTrackChanges,
  loadPreferences,
  savePreferences,
  onMount,
}: AppProps): React.JSX.Element {
  const {
    shellState,
    viewState,
    isIdle,
    isHeroError,
    isHeroProgress,
    handleOpenSettings,
    handleCloseSettings,
  } = useApp();
  const classes = useAppClasses();
  const selectionPreview = useSelectionPreview();

  React.useEffect(() => {
    onMount?.();
  }, [onMount]);

  if (viewState.view === "settings") {
    return (
      <main id="app-body">
        <SettingsView
          loadPreferences={loadPreferences}
          onBack={handleCloseSettings}
          savePreferences={savePreferences}
        />
      </main>
    );
  }

  const analyzeSection = (
    <AnalyzeSection isLoading={shellState.isAnalyzeLoading} onAnalyze={onAnalyze} />
  );
  const selectionPreviewSection = <SelectionPreview {...selectionPreview} />;

  return (
    <main id="app-body">
      <div className={classes.workflow}>
        {isIdle && (
          <HeroEmptyState>
            {analyzeSection}
            {selectionPreviewSection}
          </HeroEmptyState>
        )}
        {isHeroError && (
          <AnalysisErrorState
            error={shellState.analysisError}
            onRetryAnalysis={onRetryAnalysis}
            onRetryAnalysisQuery={onRetryAnalysisQuery}
          >
            {selectionPreviewSection}
          </AnalysisErrorState>
        )}
        {isHeroProgress && (
          <AnalysisProgressHero
            onCancelAnalysis={onCancelAnalysis}
            onRetryAnalysisQuery={onRetryAnalysisQuery}
            progress={shellState.progress}
          />
        )}
        {!isIdle && !isHeroError && !isHeroProgress && (
          <>
            {analyzeSection}
            {selectionPreviewSection}
            <ProgressPanel
              onCancelAnalysis={onCancelAnalysis}
              onRetryAnalysisQuery={onRetryAnalysisQuery}
              progress={shellState.progress}
            />
            <ResultsPanel />
          </>
        )}
        <CleanupSection
          isLoading={shellState.isCleanupLoading}
          isVisible={shellState.cleanupVisible}
          onCleanup={onCleanup}
        />
        <DisableTrackChangesSection
          isLoading={shellState.isDisableTrackChangesLoading}
          isVisible={shellState.disableTrackChangesCtaVisible}
          onDisableTrackChanges={onDisableTrackChanges}
        />
        <StatusBar status={shellState.status} />
      </div>
      <div className={classes.toolbar}>
        <SettingsToolbar onOpenSettings={handleOpenSettings} />
      </div>
    </main>
  );
}
