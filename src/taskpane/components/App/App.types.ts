export type AnalysisProfileOption = Readonly<{
  value: string;
  label: string;
}>;

export type AppProps = Readonly<{
  onMount?: () => void;
  title?: string;
  subtitle?: string;
}>;
