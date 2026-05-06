export type AppProps = Readonly<{
  onAnalyze: () => Promise<void> | void;
  onCleanup: () => Promise<void> | void;
  onDisableTrackChanges: () => Promise<void> | void;
  onSignIn: () => Promise<void> | void;
  onSignOut: () => Promise<void> | void;
  onMount?: () => void;
}>;

export type AppClasses = Readonly<{
  workflow: string;
  toolbar: string;
}>;
