import type * as React from "react";
import type { TaskpaneAnalysisErrorState } from "../../TaskpaneShellStore.types";

/** Props required to render the hero-style analysis-error surface. */
export type AnalysisErrorStateProps = Readonly<{
  error: TaskpaneAnalysisErrorState;
  onRetryAnalysis: () => Promise<void> | void;
  onRetryAnalysisQuery: () => Promise<void> | void;
  children?: React.ReactNode;
}>;

/** Griffel class slots consumed by the analysis-error surface. */
export type AnalysisErrorStateClasses = Readonly<{
  root: string;
  illustrationWrapper: string;
  illustration: string;
  illustrationDoc: string;
  illustrationAlert: string;
  copy: string;
  title: string;
  message: string;
  guidance: string;
  actions: string;
  retryButton: string;
}>;
