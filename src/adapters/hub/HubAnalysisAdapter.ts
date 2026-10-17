/**
 * Hub Analysis Adapter — implements `IAnalysisPort` against Stylistic Hub API v1.
 *
 * One analysis covers every chunk of a request: it is started with
 * `POST /documents/{documentId}/analyses`, read with `GET /analyses/{id}`, and
 * canceled as a whole. Chunk ids are the chunk indexes as strings, so the
 * hub's per-chunk results map back to the add-in's chunks. Suggestion ids are
 * issued by the hub and stay unique for good.
 *
 * Every refusal is translated into `AnalysisRequestError` carrying the hub's
 * problem `detail`, so the taskpane can show the real cause.
 *
 * @module HubAnalysisAdapter
 */

import type {
  AnalysisReadiness,
  AnalysisRequest,
  AnalysisSnapshot,
  ChunkAnalysisResult,
} from "../../domain/analysis/Analysis.types";
import { AnalysisRequestError } from "../../domain/analysis/AnalysisRequestError";
import type { IAnalysisPort } from "../../domain/ports";
import type { Suggestion } from "../../domain/suggestion/Suggestion.types";
import type { HttpClient } from "../../infrastructure/http/HttpClient";
import type {
  HubAnalysis,
  HubAnalysisRequestBody,
  HubChunk,
  HubStatus,
  HubSuggestion,
} from "./HubAnalysisAdapter.types";
import { readHubErrorDetail, readHubErrorStatus } from "./HubProblemParser.helpers";

/** Talks to the hub's analysis resources on behalf of the pipeline. */
export class HubAnalysisAdapter implements IAnalysisPort {
  constructor(private readonly http: HttpClient) {}

  /** Asks the hub whether it can analyze; an unreachable hub is not ready. */
  async getReadiness(): Promise<AnalysisReadiness> {
    try {
      const status = await this.http.get<HubStatus>("/status");
      return status.analysis.ready
        ? { ready: true }
        : { ready: false, reason: status.analysis.reason };
    } catch (error) {
      return {
        ready: false,
        reason: `No se pudo conectar con el hub: ${readHubErrorDetail(error)}`,
      };
    }
  }

  /** Starts one analysis for every chunk of the request. */
  async startAnalysis(request: AnalysisRequest): Promise<AnalysisSnapshot> {
    const body: HubAnalysisRequestBody = {
      genre: request.genre,
      chunks: request.chunks.map((chunk) => ({ id: String(chunk.index), text: chunk.text })),
    };
    const path = `/documents/${encodeURIComponent(request.documentUuid)}/analyses`;

    return this.toSnapshot(await this.call(() => this.http.post<HubAnalysis>(path, body)));
  }

  /** Reads the current state of an analysis. */
  async getAnalysis(analysisId: string): Promise<AnalysisSnapshot> {
    const path = `/analyses/${encodeURIComponent(analysisId)}`;
    return this.toSnapshot(await this.call(() => this.http.get<HubAnalysis>(path)));
  }

  /** Cancels an analysis and every chunk still running in it. */
  async cancelAnalysis(analysisId: string): Promise<void> {
    const path = `/analyses/${encodeURIComponent(analysisId)}/cancel`;
    await this.call(() => this.http.post<HubAnalysis>(path, undefined));
  }

  /** Runs a hub request, translating any failure into `AnalysisRequestError`. */
  private async call<T>(request: () => Promise<T>): Promise<T> {
    try {
      return await request();
    } catch (error) {
      throw new AnalysisRequestError(readHubErrorDetail(error), readHubErrorStatus(error));
    }
  }

  /** Maps a hub analysis to the domain snapshot. */
  private toSnapshot(analysis: HubAnalysis): AnalysisSnapshot {
    return {
      analysisId: analysis.id,
      status: analysis.status,
      chunks: analysis.chunks.map((chunk) => this.toChunkResult(chunk)),
    };
  }

  /** Maps one hub chunk back to the add-in chunk whose index is its id. */
  private toChunkResult(chunk: HubChunk): ChunkAnalysisResult {
    const result: ChunkAnalysisResult = {
      chunkIndex: Number.parseInt(chunk.id, 10),
      status: chunk.status,
      suggestions: chunk.suggestions.map((suggestion) => this.toSuggestion(suggestion)),
      dropped: chunk.dropped,
    };

    if (chunk.error !== undefined) {
      result.error = chunk.error;
    }

    return result;
  }

  /** Maps one hub suggestion to the domain suggestion. */
  private toSuggestion(raw: HubSuggestion): Suggestion {
    const suggestion: Suggestion = {
      id: raw.id,
      context: raw.context,
      anchor: raw.anchor,
      justification: raw.justification,
      category: raw.category,
      severity: raw.severity,
      type: raw.type,
    };

    if (raw.type === "track-change") {
      suggestion.suggestedText = raw.suggestedText;
    }

    return suggestion;
  }
}
