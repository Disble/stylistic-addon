/* global console */

/**
 * Text chunking module — splits document text into sized chunks at paragraph
 * boundaries for sequential delivery to the Mastra editorial workflow.
 *
 * Design constraints:
 * - Never splits mid-paragraph (preserves semantic context for the AI).
 * - Paragraphs exceeding `maxChunkSize` are sent as-is (backend must handle).
 * - Small documents (< maxChunkSize) produce a single chunk.
 * - Pure function with no side effects — depends only on domain types.
 *
 * @module chunker
 */

import type { TextChunk } from "../domain/chunking/TextChunk.types";
import { DEFAULT_MAX_CHUNK_SIZE } from "./config";

/** Regex that matches paragraph separators: one or more blank lines. */
const PARAGRAPH_SEPARATOR = /\r?\n\s*\r?\n/;

/**
 * Splits a document's full text into chunks that respect paragraph boundaries.
 *
 * Paragraphs are accumulated into a chunk until adding the next paragraph
 * would exceed `maxChunkSize`. At that point, the current chunk is finalized
 * and a new one begins.
 *
 * @param text         - The full document text.
 * @param maxChunkSize - Maximum characters per chunk. Defaults to {@link DEFAULT_MAX_CHUNK_SIZE}.
 * @returns An array of {@link TextChunk} objects with positional metadata.
 *          Returns an empty array if `text` is empty or whitespace-only.
 */
export function splitText(
  text: string,
  maxChunkSize: number = DEFAULT_MAX_CHUNK_SIZE,
): TextChunk[] {
  console.log(
    `✂️ [Chunker] splitText: ${text.length} chars, maxChunkSize: ${maxChunkSize}`,
  );
  if (!text || text.trim().length === 0) {
    return [];
  }

  const paragraphs = text.split(PARAGRAPH_SEPARATOR);
  const chunks: TextChunk[] = [];
  let currentParts: string[] = [];
  let currentLength = 0;
  let currentOffset = 0;

  for (const paragraph of paragraphs) {
    const separatorLength = currentLength > 0 ? 2 : 0;
    const wouldExceed =
      currentLength > 0 &&
      currentLength + separatorLength + paragraph.length > maxChunkSize;

    if (wouldExceed) {
      chunks.push({
        text: currentParts.join("\n\n"),
        index: chunks.length,
        total: 0,
        startOffset: currentOffset,
      });

      currentOffset += currentLength;
      currentOffset += findSeparatorLength(text, currentOffset);
      currentParts = [paragraph];
      currentLength = paragraph.length;
    } else {
      if (currentParts.length > 0) {
        currentLength += separatorLength;
      }
      currentParts.push(paragraph);
      currentLength += paragraph.length;
    }
  }

  if (currentParts.length > 0) {
    chunks.push({
      text: currentParts.join("\n\n"),
      index: chunks.length,
      total: 0,
      startOffset: currentOffset,
    });
  }

  for (const chunk of chunks) {
    chunk.total = chunks.length;
  }

  console.log(`✂️ [Chunker] ${chunks.length} chunk(s) generados`);
  return chunks;
}

/**
 * Finds the length of the paragraph separator at a given position in the
 * original text. Accounts for `\r\n` vs `\n` line endings and multiple
 * blank lines.
 */
function findSeparatorLength(text: string, offset: number): number {
  let length = 0;
  let i = offset;
  while (i < text.length) {
    if (
      text[i] === "\r" ||
      text[i] === "\n" ||
      text[i] === " " ||
      text[i] === "\t"
    ) {
      length++;
      i++;
    } else {
      break;
    }
  }
  return length;
}
