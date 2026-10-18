import type { AnalysisStatus, ChunkAnalysisStatus } from "../../domain/analysis/Analysis.types";
import type { AnalysisProfileId } from "../../domain/Profile.types";
import type { SuggestionSeverity, SuggestionType } from "../../domain/suggestion/Suggestion.types";

/** `GET /status` response. */
export interface HubStatus {
  version: string;
  analysis: {
    ready: boolean;
    reason?: string;
  };
}

/** `POST /documents/{documentId}/analyses` request body. */
export interface HubAnalysisRequestBody {
  genre: AnalysisProfileId;
  chunks: Array<{ id: string; text: string }>;
}

/** One suggestion as the hub issues it. */
export interface HubSuggestion {
  id: string;
  context: string;
  anchor: string;
  justification: string;
  category: string;
  severity: SuggestionSeverity;
  type: SuggestionType;
  suggestedText?: string;
}

/** One of the add-in's chunks within a hub analysis. */
export interface HubChunk {
  id: string;
  status: ChunkAnalysisStatus;
  suggestions: HubSuggestion[];
  dropped: number;
  error?: string;
}

/** An analysis as the hub returns it from start, read and cancel. */
export interface HubAnalysis {
  id: string;
  documentId: string;
  status: AnalysisStatus;
  chunks: HubChunk[];
}
