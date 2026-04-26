/** Paragraph snapshot required to reconstruct structured document text. */
export type ParagraphSnapshot = {
  text?: string;
  styleBuiltIn?: string;
  firstLineIndent?: number;
  leftIndent?: number;
};
