/** Props required to render the resolved-comment cleanup section. */
export type CleanupSectionProps = Readonly<{
  isLoading: boolean;
  isVisible: boolean;
  onCleanup: () => Promise<void> | void;
}>;

/** Griffel class slots consumed by the cleanup section. */
export type CleanupSectionClasses = Readonly<{
  root: string;
  button: string;
}>;
