import type * as React from "react";
import type { TaskpaneAnalysisErrorState } from "../../TaskpaneShellStore.types";

/** Props required to render the hero-style analysis-error surface. */
export type AnalysisErrorStateProps = Readonly<{
  error: TaskpaneAnalysisErrorState;
  onRetryAnalysis: () => Promise<void> | void;
  onRetryAnalysisQuery: () => Promise<void> | void;
  children?: React.ReactNode;
}>;

/** Copy contract selected from the current analysis-error retry mode. */
export type AnalysisErrorRetryPresentation = Readonly<{
  title: string;
  guidance: string;
  actionLabel: string;
}>;

/** Griffel class slots consumed by the analysis-error surface. */
export type AnalysisErrorStateClasses = Readonly<{
  illustration: string;
  illustrationDoc: string;
  illustrationAlert: string;
  guidance: string;
  retryButton: string;
}>;
