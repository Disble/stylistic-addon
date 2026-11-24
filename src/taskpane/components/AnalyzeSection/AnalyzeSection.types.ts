/** Props required to render the analyze action section. */
export type AnalyzeSectionProps = Readonly<{
  isLoading: boolean;
  onAnalyze: () => Promise<void> | void;
}>;

/** Griffel class slots consumed by the analyze section. */
export type AnalyzeSectionClasses = Readonly<{
  root: string;
  button: string;
}>;
