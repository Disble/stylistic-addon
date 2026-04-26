/** A chunk of document text prepared for sending to the Mastra workflow. */
export interface TextChunk {
  /** The text content of this chunk. */
  text: string;

  /** Zero-based index of this chunk within the full document. */
  index: number;

  /** Total number of chunks the document was split into. */
  total: number;

  /** Character offset where this chunk starts in the original document. */
  startOffset: number;
}
