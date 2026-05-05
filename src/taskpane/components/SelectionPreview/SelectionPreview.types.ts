export type SelectionPreviewClasses = Readonly<{
  root: string;
  body: string;
  quote: string;
  hint: string;
}>;

export type SelectionPreviewProps = Readonly<{
  isVisible: boolean;
  charCount: number;
  preview: string;
  classes: SelectionPreviewClasses;
}>;
