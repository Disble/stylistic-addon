export type CleanupSectionProps = Readonly<{
  isLoading: boolean;
  isVisible: boolean;
  onCleanup: () => Promise<void> | void;
}>;
