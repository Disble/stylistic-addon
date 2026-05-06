export type SelectionPreviewClasses = Readonly<{
  root: string;
  icon: string;
  label: string;
}>;

export type SelectionPreviewProps = Readonly<{
  isVisible: boolean;
  charCount: number;
  preview: string;
  classes: SelectionPreviewClasses;
}>;
