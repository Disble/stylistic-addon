export type AnalyzeSectionProps = Readonly<{
  isLoading: boolean;
  onAnalyze: () => Promise<void> | void;
}>;

export type AnalyzeSectionClasses = Readonly<{
  root: string;
  button: string;
}>;
