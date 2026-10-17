/* global RequestInit */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { AnalysisRequestError } from "../../../domain/analysis/AnalysisRequestError";
import type { AnalysisRequest } from "../../../domain/analysis/Analysis.types";
import { HttpClient } from "../../../infrastructure/http/HttpClient";
import { HubAnalysisAdapter } from "../HubAnalysisAdapter";

type FetchMock = ReturnType<typeof vi.fn> & typeof fetch;

const BASE_URL = "https://localhost:3000/api/v1";
const DOCUMENT_UUID = "11111111-1111-4111-8111-111111111111";

function jsonResponse(status: number, body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });
}

function problemResponse(status: number, code: string, detail: string): Response {
  return new Response(JSON.stringify({ type: "about:blank", title: code, status, detail, code }), {
    status,
    headers: { "content-type": "application/problem+json" },
  });
}

function hubAnalysis(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    id: "8e0f8a52-6c1f-4a7e-9d43-1b6b1f1d2c3a",
    documentId: DOCUMENT_UUID,
    status: "completed",
    chunks: [
      {
        id: "0",
        status: "completed",
        dropped: 1,
        suggestions: [
          {
            id: "0123456789abcdef",
            context: "Era una vez un texto.",
            anchor: "texto",
            justification: "Más preciso.",
            category: "lexico",
            severity: "low",
            type: "track-change",
            suggestedText: "relato",
          },
          {
            id: "fedcba9876543210",
            context: "Era una vez un texto.",
            anchor: "Era una vez",
            justification: "Fórmula gastada.",
            category: "estilo",
            severity: "medium",
            type: "comment-only",
          },
        ],
      },
      {
        id: "1",
        status: "failed",
        dropped: 0,
        suggestions: [],
        error: "The model request failed: 400 MissingSessionID",
      },
    ],
    profile: "pending",
    createdAt: "2026-10-17T19:54:00.000Z",
    ...overrides,
  };
}

function makeRequest(): AnalysisRequest {
  return {
    documentUuid: DOCUMENT_UUID,
    genre: "narrativa-literaria",
    chunks: [
      { text: "Era una vez un texto.", index: 0, total: 2, startOffset: 0 },
      { text: "Segundo fragmento.", index: 1, total: 2, startOffset: 23 },
    ],
  };
}

describe("HubAnalysisAdapter", () => {
  let fetchImpl: FetchMock;
  let adapter: HubAnalysisAdapter;

  beforeEach(() => {
    fetchImpl = vi.fn() as unknown as FetchMock;
    adapter = new HubAnalysisAdapter(
      new HttpClient({ baseUrl: BASE_URL, fetchImpl, getAuthToken: () => "hub-token" })
    );
    vi.spyOn(console, "log").mockImplementation(() => {});
    vi.spyOn(console, "error").mockImplementation(() => {});
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  function lastCall(): { url: string; init: RequestInit } {
    const [url, init] = fetchImpl.mock.calls[fetchImpl.mock.calls.length - 1] as [
      string,
      RequestInit,
    ];
    return { url, init };
  }

  describe("getReadiness", () => {
    it("reports ready when the hub can analyze", async () => {
      fetchImpl.mockResolvedValueOnce(
        jsonResponse(200, { version: "0.0.1", analysis: { ready: true } })
      );

      await expect(adapter.getReadiness()).resolves.toEqual({ ready: true });
      const { url, init } = lastCall();
      expect(url).toBe(`${BASE_URL}/status`);
      expect(init.method).toBe("GET");
      expect(new Headers(init.headers).get("Authorization")).toBe("Bearer hub-token");
    });

    it("passes on the hub's reason when it cannot analyze", async () => {
      fetchImpl.mockResolvedValueOnce(
        jsonResponse(200, {
          version: "0.0.1",
          analysis: { ready: false, reason: "Choose a provider and a model in the hub." },
        })
      );

      await expect(adapter.getReadiness()).resolves.toEqual({
        ready: false,
        reason: "Choose a provider and a model in the hub.",
      });
    });

    it("reports not ready with the problem detail when the hub refuses", async () => {
      fetchImpl.mockResolvedValueOnce(problemResponse(401, "unauthorized", "Missing token"));

      await expect(adapter.getReadiness()).resolves.toEqual({
        ready: false,
        reason: "No se pudo conectar con el hub: Missing token",
      });
    });

    it("reports not ready when the hub is unreachable", async () => {
      fetchImpl.mockRejectedValueOnce(new TypeError("Failed to fetch"));

      await expect(adapter.getReadiness()).resolves.toEqual({
        ready: false,
        reason: "No se pudo conectar con el hub: Failed to fetch",
      });
    });
  });

  describe("startAnalysis", () => {
    it("posts every chunk with its index as id and maps the hub analysis", async () => {
      fetchImpl.mockResolvedValueOnce(jsonResponse(202, hubAnalysis({ status: "partial" })));

      const snapshot = await adapter.startAnalysis(makeRequest());

      const { url, init } = lastCall();
      expect(url).toBe(`${BASE_URL}/documents/${DOCUMENT_UUID}/analyses`);
      expect(init.method).toBe("POST");
      expect(JSON.parse(init.body as string)).toEqual({
        genre: "narrativa-literaria",
        chunks: [
          { id: "0", text: "Era una vez un texto." },
          { id: "1", text: "Segundo fragmento." },
        ],
      });
      expect(snapshot).toEqual({
        analysisId: "8e0f8a52-6c1f-4a7e-9d43-1b6b1f1d2c3a",
        status: "partial",
        chunks: [
          {
            chunkIndex: 0,
            status: "completed",
            dropped: 1,
            suggestions: [
              {
                id: "0123456789abcdef",
                context: "Era una vez un texto.",
                anchor: "texto",
                justification: "Más preciso.",
                category: "lexico",
                severity: "low",
                type: "track-change",
                suggestedText: "relato",
              },
              {
                id: "fedcba9876543210",
                context: "Era una vez un texto.",
                anchor: "Era una vez",
                justification: "Fórmula gastada.",
                category: "estilo",
                severity: "medium",
                type: "comment-only",
              },
            ],
          },
          {
            chunkIndex: 1,
            status: "failed",
            dropped: 0,
            suggestions: [],
            error: "The model request failed: 400 MissingSessionID",
          },
        ],
      });
    });

    it("throws the problem detail and status when the hub refuses", async () => {
      fetchImpl.mockResolvedValueOnce(
        problemResponse(409, "no_model", "Choose a model in the hub before analyzing.")
      );

      const error = await adapter.startAnalysis(makeRequest()).catch((caught: unknown) => caught);

      expect(error).toBeInstanceOf(AnalysisRequestError);
      expect((error as AnalysisRequestError).message).toBe(
        "Choose a model in the hub before analyzing."
      );
      expect((error as AnalysisRequestError).status).toBe(409);
    });
  });

  describe("getAnalysis", () => {
    it("reads the analysis by id", async () => {
      fetchImpl.mockResolvedValueOnce(jsonResponse(200, hubAnalysis({ status: "running" })));

      const snapshot = await adapter.getAnalysis("8e0f8a52-6c1f-4a7e-9d43-1b6b1f1d2c3a");

      expect(lastCall().url).toBe(`${BASE_URL}/analyses/8e0f8a52-6c1f-4a7e-9d43-1b6b1f1d2c3a`);
      expect(lastCall().init.method).toBe("GET");
      expect(snapshot.status).toBe("running");
    });

    it("throws a status-less error carrying the cause when the hub is unreachable", async () => {
      fetchImpl.mockRejectedValueOnce(new TypeError("Failed to fetch"));

      const error = await adapter.getAnalysis("a1").catch((caught: unknown) => caught);

      expect(error).toBeInstanceOf(AnalysisRequestError);
      expect((error as AnalysisRequestError).message).toBe("Failed to fetch");
      expect((error as AnalysisRequestError).status).toBeUndefined();
    });
  });

  describe("cancelAnalysis", () => {
    it("cancels the whole analysis", async () => {
      fetchImpl.mockResolvedValueOnce(jsonResponse(200, hubAnalysis({ status: "canceled" })));

      await adapter.cancelAnalysis("a1");

      expect(lastCall().url).toBe(`${BASE_URL}/analyses/a1/cancel`);
      expect(lastCall().init.method).toBe("POST");
      expect(lastCall().init.body).toBeUndefined();
    });

    it("throws the problem detail when the hub refuses", async () => {
      fetchImpl.mockResolvedValueOnce(problemResponse(404, "not_found", "No such analysis"));

      await expect(adapter.cancelAnalysis("a1")).rejects.toThrow("No such analysis");
    });
  });
});
