import * as React from "react";
import { AnalyzeSection } from "../AnalyzeSection";
import { AuthSection } from "../AuthSection";
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
import { useApp, useAppClasses } from "./useApp";

/**
 * React shell for the taskpane.
 *
 * Renders one of three top-level surfaces based on auth + view state:
 * - unauthenticated → `AuthSection` (login screen, no toolbar, no settings access)
 * - authenticated + `view === "settings"` → `SettingsView` (account + analysis-profile + future setting groups)
 * - authenticated + `view === "main"` → analysis workflow + persistent `SettingsToolbar`
 *
 * Preserves the legacy `#app-body` DOM anchor so `taskpane.css` can keep owning the
 * outer flex layout (height chain + padding) without a port through Griffel.
 */
export function App({
  onAnalyze,
  onCleanup,
  onDisableTrackChanges,
  onSignIn,
  onSignOut,
  onMount,
}: AppProps): React.JSX.Element {
  const {
    authState,
    shellState,
    viewState,
    isIdle,
    handleOpenSettings,
    handleCloseSettings,
  } = useApp();
  const classes = useAppClasses();
  const selectionPreview = useSelectionPreview();

  React.useEffect(() => {
    onMount?.();
  }, [onMount]);

  if (authState.status !== "authenticated") {
    return (
      <main id="app-body">
        <AuthSection
          error={authState.error}
          isSigningIn={authState.isSigningIn}
          onSignIn={onSignIn}
          status={authState.status}
        />
      </main>
    );
  }

  if (viewState.view === "settings") {
    return (
      <main id="app-body">
        <SettingsView
          isSigningOut={authState.isSigningOut}
          onBack={handleCloseSettings}
          onSignOut={onSignOut}
          session={authState.session}
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
        {isIdle ? (
          <HeroEmptyState>
            {analyzeSection}
            {selectionPreviewSection}
          </HeroEmptyState>
        ) : (
          <>
            {analyzeSection}
            {selectionPreviewSection}
            <ProgressPanel progress={shellState.progress} />
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
