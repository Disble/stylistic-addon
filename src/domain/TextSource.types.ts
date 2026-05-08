/**
 * Result of the text-source resolution step at the start of the analysis
 * pipeline.
 */
export interface TextSource {
  /** The plain text to analyze (selection or full document). */
  text: string;

  /** Whether the text was read from the user's current selection. */
  isSelection: boolean;

  /** Stable add-in-generated UUID persisted in the active Word document. */
  documentUuid: string;
}
