import type { Suggestion } from "../suggestion/Suggestion.types";
import type { AnalysisSnapshot, ChunkAnalysisResult } from "./Analysis.types";

/** Returns every suggestion of an analysis, in chunk order. */
export function collectAnalysisSuggestions(snapshot: AnalysisSnapshot): Suggestion[] {
  return sortByChunkIndex(snapshot.chunks).flatMap((chunk) => chunk.suggestions);
}

/**
 * Describes each failed chunk with the backend's real cause, prefixed with its
 * position when the analysis has more than one chunk.
 */
export function describeAnalysisFailures(snapshot: AnalysisSnapshot): string[] {
  const chunkCount = snapshot.chunks.length;

  return sortByChunkIndex(snapshot.chunks)
    .filter((chunk) => chunk.status === "failed")
    .map((chunk) => {
      const cause = chunk.error ?? "el backend no informó la causa";
      return chunkCount > 1 ? `Fragmento ${chunk.chunkIndex + 1}: ${cause}` : cause;
    });
}

/** Counts the chunks that reached a terminal status. */
export function countFinishedChunks(snapshot: AnalysisSnapshot): number {
  return snapshot.chunks.filter((chunk) => chunk.status !== "queued" && chunk.status !== "running")
    .length;
}

/** Counts the suggestions the backend discarded across every chunk. */
export function countDroppedSuggestions(snapshot: AnalysisSnapshot): number {
  return snapshot.chunks.reduce((total, chunk) => total + chunk.dropped, 0);
}

/** Explains why an analysis produced no suggestions. */
export function describeNoSuggestionsReason(failures: string[]): string {
  return failures.length > 0
    ? `El análisis falló: ${failures.join(" · ")}`
    : "No se encontraron sugerencias editoriales.";
}

/** Returns a copy of the chunks ordered by chunk index. */
function sortByChunkIndex(chunks: ChunkAnalysisResult[]): ChunkAnalysisResult[] {
  return [...chunks].sort((a, b) => a.chunkIndex - b.chunkIndex);
}
