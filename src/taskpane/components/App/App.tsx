import * as React from "react";
import { AnalysisProfileSection } from "../AnalysisProfileSection";
import { AnalyzeSection } from "../AnalyzeSection";
import { CleanupSection } from "../CleanupSection";
import { DisableTrackChangesSection } from "../DisableTrackChangesSection";
import { ProgressPanel } from "../ProgressPanel";
import { ResultsPanel } from "../ResultsPanel";
import { StatusBar } from "../StatusBar";
import { TaskpaneHeader } from "../TaskpaneHeader";
import {
  ANALYSIS_PROFILE_OPTIONS,
  DEFAULT_APP_SUBTITLE,
  DEFAULT_APP_TITLE,
} from "./App.constants";
import type { AppProps } from "./App.types";
import { useApp } from "./useApp";

/**
 * React shell for the taskpane.
 * It preserves the existing DOM anchor IDs so the legacy composition root can bind safely during migration.
 */
export function App({
  onMount,
  title = DEFAULT_APP_TITLE,
  subtitle = DEFAULT_APP_SUBTITLE,
}: AppProps): React.JSX.Element {
  const shellState = useApp();

  React.useEffect(() => {
    onMount?.();
  }, [onMount]);

  return (
    <main id="app-body">
      <TaskpaneHeader title={title} subtitle={subtitle} />
      <AnalysisProfileSection
        isDisabled={shellState.isAnalyzeLoading}
        options={ANALYSIS_PROFILE_OPTIONS}
      />
      <AnalyzeSection isLoading={shellState.isAnalyzeLoading} />
      <ProgressPanel progress={shellState.progress} />
      <ResultsPanel />
      <CleanupSection isVisible={shellState.cleanupVisible} />
      <DisableTrackChangesSection isVisible={shellState.disableTrackChangesCtaVisible} />
      <StatusBar status={shellState.status} />
    </main>
  );
}
