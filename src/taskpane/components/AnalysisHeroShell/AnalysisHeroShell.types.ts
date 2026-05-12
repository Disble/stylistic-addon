import type * as React from "react";

/** Props required to render the shared hero shell used by analysis states. */
export type AnalysisHeroShellProps = Readonly<{
  sectionLabelledById: string;
  testId: string;
  title: string;
  message: string;
  illustration: React.ReactNode;
  messageSupplement?: React.ReactNode;
  actions?: React.ReactNode;
  dataRetryKind?: string;
}>;

/** Griffel class slots consumed by the shared analysis hero shell. */
export type AnalysisHeroShellClasses = Readonly<{
  root: string;
  illustrationWrapper: string;
  copy: string;
  title: string;
  message: string;
  actions: string;
}>;
