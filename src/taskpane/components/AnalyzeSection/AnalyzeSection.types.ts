export type AnalyzeSectionProps = Readonly<{
  isLoading: boolean;
  onAnalyze: () => Promise<void> | void;
}>;
