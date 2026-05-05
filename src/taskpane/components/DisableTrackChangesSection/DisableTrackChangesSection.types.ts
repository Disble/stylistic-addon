export type DisableTrackChangesSectionProps = Readonly<{
  isLoading: boolean;
  isVisible: boolean;
  onDisableTrackChanges: () => Promise<void> | void;
}>;

export type DisableTrackChangesSectionClasses = Readonly<{
  root: string;
  button: string;
}>;
