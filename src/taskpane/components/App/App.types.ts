export type AnalysisProfileOption = Readonly<{
  value: string;
  label: string;
}>;

export type AppProps = Readonly<{
  onAnalyze: () => Promise<void> | void;
  onCleanup: () => Promise<void> | void;
  onDisableTrackChanges: () => Promise<void> | void;
  onMount?: () => void;
}>;
