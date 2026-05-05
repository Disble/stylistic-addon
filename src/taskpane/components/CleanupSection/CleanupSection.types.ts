export type CleanupSectionProps = Readonly<{
  isLoading: boolean;
  isVisible: boolean;
  onCleanup: () => Promise<void> | void;
}>;

export type CleanupSectionClasses = Readonly<{
  root: string;
  button: string;
}>;
