import type { TextChunk } from "../chunking/TextChunk.types";
import type { AnalysisProfileId } from "../Profile.types";
import type { Suggestion } from "../suggestion/Suggestion.types";

/** Whether the analysis backend can analyze right now. */
export interface AnalysisReadiness {
  /** True when the backend is reachable and has a model to analyze with. */
  ready: boolean;

  /** User-facing reason shown when `ready` is false. */
  reason?: string;
}

/** Everything the backend needs to start analyzing one document. */
export interface AnalysisRequest {
  /** Stable document UUID persisted by the add-in in the Word document. */
  documentUuid: string;

  /** Editorial genre the document is analyzed for. */
  genre: AnalysisProfileId;

  /** Text chunks to analyze, in document order. */
  chunks: TextChunk[];
}

/** Lifecycle of a whole analysis. Every status except `running` is terminal. */
export type AnalysisStatus = "running" | "completed" | "partial" | "failed" | "canceled";

/** Lifecycle of one chunk inside an analysis. */
export type ChunkAnalysisStatus = "queued" | "running" | "completed" | "failed" | "canceled";

/** Current outcome of one chunk inside an analysis. */
export interface ChunkAnalysisResult {
  /** Zero-based index of the chunk this result belongs to. */
  chunkIndex: number;

  /** Current lifecycle status of the chunk. */
  status: ChunkAnalysisStatus;

  /** Suggestions produced for the chunk; empty until it completes. */
  suggestions: Suggestion[];

  /** Suggestions the backend discarded because they broke the contract. */
  dropped: number;

  /** Real cause reported by the backend when the chunk failed. */
  error?: string;
}

/** Point-in-time view of an analysis and each of its chunks. */
export interface AnalysisSnapshot {
  /** Backend-issued analysis identifier used to poll and cancel. */
  analysisId: string;

  /** Current lifecycle status of the whole analysis. */
  status: AnalysisStatus;

  /** Per-chunk outcomes, ordered by chunk index. */
  chunks: ChunkAnalysisResult[];
}

/** Analysis references the taskpane needs for cancel and retry-query actions. */
export interface AnalysisSessionSnapshot {
  /** Whether the originating analysis targeted the current selection. */
  isSelection: boolean;

  /** Analysis still being polled, which the user can cancel. */
  activeAnalysisId: string | null;

  /** Analysis whose polling failed locally and can be queried again. */
  retryableAnalysisId: string | null;
}
