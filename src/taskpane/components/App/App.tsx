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

/**
 * React shell for the taskpane.
 * It preserves the existing DOM anchor IDs so the legacy composition root can bind safely during migration.
 */
export function App({
  title = DEFAULT_APP_TITLE,
  subtitle = DEFAULT_APP_SUBTITLE,
}: AppProps): React.JSX.Element {
  return (
    <>
      <section id="sideload-msg" className="sideload-msg">
        <h2>Please sideload your add-in to see app body.</h2>
      </section>

      <main id="app-body" style={{ display: "none" }}>
        <TaskpaneHeader title={title} subtitle={subtitle} />
        <AnalysisProfileSection options={ANALYSIS_PROFILE_OPTIONS} />
        <AnalyzeSection />
        <ProgressPanel />
        <ResultsPanel />
        <CleanupSection />
        <DisableTrackChangesSection />
        <StatusBar />
      </main>
    </>
  );
}
