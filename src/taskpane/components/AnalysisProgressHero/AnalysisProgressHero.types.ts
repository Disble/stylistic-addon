import type { TaskpaneShellProgress } from "../../TaskpaneShellStore.types";

/** Props required to render the hero-style analysis-in-progress surface. */
export type AnalysisProgressHeroProps = Readonly<{
  progress: TaskpaneShellProgress;
  onCancelAnalysis: () => Promise<void> | void;
  onRetryAnalysisQuery: () => Promise<void> | void;
}>;

/** Griffel class slots consumed by the analysis-progress hero. */
export type AnalysisProgressHeroClasses = Readonly<{
  illustration: string;
  illustrationDoc: string;
  illustrationWand: string;
  illustrationSparkle: string;
  illustrationSparkle1: string;
  illustrationSparkle2: string;
  illustrationSparkle3: string;
  illustrationLine: string;
  illustrationLineDelay1: string;
  illustrationLineDelay2: string;
  illustrationLineDelay3: string;
  illustrationLineDelay4: string;
  illustrationLineDelay5: string;
  illustrationLineDelay6: string;
  progressField: string;
  primaryButton: string;
}>;
