/** Props required to render the disable-track-changes call to action. */
export type DisableTrackChangesSectionProps = Readonly<{
  isLoading: boolean;
  isVisible: boolean;
  onDisableTrackChanges: () => Promise<void> | void;
}>;

/** Griffel class slots consumed by the disable-track-changes section. */
export type DisableTrackChangesSectionClasses = Readonly<{
  root: string;
  button: string;
}>;
