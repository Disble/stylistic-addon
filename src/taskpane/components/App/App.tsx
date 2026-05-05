import * as React from "react";
import { AnalysisProfileSection } from "../AnalysisProfileSection";
import { AnalyzeSection } from "../AnalyzeSection";
import { AuthSection } from "../AuthSection";
import { CleanupSection } from "../CleanupSection";
import { DisableTrackChangesSection } from "../DisableTrackChangesSection";
import { ProgressPanel } from "../ProgressPanel";
import { ResultsPanel } from "../ResultsPanel";
import { SelectionPreview, useSelectionPreview } from "../SelectionPreview";
import { StatusBar } from "../StatusBar";
import { ANALYSIS_PROFILE_OPTIONS } from "./App.constants";
import type { AppProps } from "./App.types";
import { useApp } from "./useApp";

/**
 * React shell for the taskpane.
 * It preserves the existing DOM anchor IDs so the legacy composition root can bind safely during migration.
 */
export function App({
  onAnalyze,
  onCleanup,
  onDisableTrackChanges,
  onSignIn,
  onSignOut,
  onMount,
}: AppProps): React.JSX.Element {
  const { authState, handleGeneroChange, shellState } = useApp();
  const selectionPreview = useSelectionPreview();

  React.useEffect(() => {
    onMount?.();
  }, [onMount]);

  return (
    <main id="app-body">
      <AuthSection
        error={authState.error}
        isSigningIn={authState.isSigningIn}
        isSigningOut={authState.isSigningOut}
        onSignIn={onSignIn}
        onSignOut={onSignOut}
        session={authState.session}
        status={authState.status}
      />
      {authState.status !== "authenticated" ? null : (
        <>
      <AnalysisProfileSection
        isDisabled={shellState.isAnalyzeLoading}
        onGeneroChange={handleGeneroChange}
        options={ANALYSIS_PROFILE_OPTIONS}
        selectedGenero={shellState.selectedGenero}
      />
      <AnalyzeSection isLoading={shellState.isAnalyzeLoading} onAnalyze={onAnalyze} />
      <SelectionPreview {...selectionPreview} />
      <ProgressPanel progress={shellState.progress} />
      <ResultsPanel />
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
        </>
      )}
    </main>
  );
}
