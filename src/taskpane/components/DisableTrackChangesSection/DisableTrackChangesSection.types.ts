export type DisableTrackChangesSectionProps = Readonly<{
  isLoading: boolean;
  isVisible: boolean;
  onDisableTrackChanges: () => Promise<void> | void;
}>;
