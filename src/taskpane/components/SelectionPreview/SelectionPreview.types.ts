/** Griffel class slots consumed by the selection preview banner. */
export type SelectionPreviewClasses = Readonly<{
  root: string;
  icon: string;
  label: string;
}>;

/** Props required to render the current Word selection preview. */
export type SelectionPreviewProps = Readonly<{
  isVisible: boolean;
  charCount: number;
  preview: string;
  classes: SelectionPreviewClasses;
}>;
