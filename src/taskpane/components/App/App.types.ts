/** Root taskpane application callbacks provided by the composition layer. */
export type AppProps = Readonly<{
  onAnalyze: () => Promise<void> | void;
  onCleanup: () => Promise<void> | void;
  onDisableTrackChanges: () => Promise<void> | void;
  onSignIn: () => Promise<void> | void;
  onSignOut: () => Promise<void> | void;
  onMount?: () => void;
}>;

/** Griffel class slots consumed by the App component. */
export type AppClasses = Readonly<{
  workflow: string;
  toolbar: string;
}>;
