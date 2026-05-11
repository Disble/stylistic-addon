import * as React from "react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { createRoot } from "react-dom/client";
import {
  bootstrapTaskpane,
  handleAnalyze,
  handleCancelAnalysis,
  handleCleanup,
  handleDisableTrackChanges,
  handleLoadPreferences,
  handleRetryAnalysisQuery,
  handleSavePreferences,
  handleSignIn,
  handleSignOut,
} from "./taskpane";
import { App } from "./components/App";

/**
 * React bootstrap entrypoint for the Stylistic taskpane.
 * Follows the Office React template bootstrap pattern and lets the mounted React shell
 * trigger taskpane DOM binding once its anchor elements exist.
 */
function renderTaskpaneShell(): void {
  const office = globalThis.Office;
  const rootElement = document.getElementById("container");

  if (!office?.onReady) {
    return;
  }

  const root = rootElement ? createRoot(rootElement) : undefined;

  office.onReady((info) => {
    const wordHost = office.HostType?.Word ?? "Word";
    if (info.host !== wordHost) {
      return;
    }

    root?.render(
      React.createElement(
        FluentProvider,
        { theme: webLightTheme },
        React.createElement(App, {
          onAnalyze: handleAnalyze,
          onRetryAnalysis: handleAnalyze,
          onCancelAnalysis: handleCancelAnalysis,
          onRetryAnalysisQuery: handleRetryAnalysisQuery,
          onCleanup: handleCleanup,
          onDisableTrackChanges: handleDisableTrackChanges,
          onSignIn: handleSignIn,
          onSignOut: handleSignOut,
          loadPreferences: handleLoadPreferences,
          savePreferences: handleSavePreferences,
          onMount: () => bootstrapTaskpane(),
        })
      )
    );
  });
}

renderTaskpaneShell();
