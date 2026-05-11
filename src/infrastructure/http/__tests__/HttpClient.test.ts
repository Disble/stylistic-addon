import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { HttpClient } from "../HttpClient";
import { HttpError } from "../HttpError";

type FetchMock = ReturnType<typeof vi.fn> & typeof fetch;

function createFetchMock(): FetchMock {
  return vi.fn() as unknown as FetchMock;
}

function makeJsonResponse(status: number, body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json" },
  });
}

function makeEmptyResponse(status: number): Response {
  return new Response(null, { status });
}

function makeTextResponse(status: number, body: string): Response {
  return new Response(body, {
    status,
    headers: { "content-type": "text/plain" },
  });
}

describe("HttpClient.get", () => {
  let fetchImpl: FetchMock;

  beforeEach(() => {
    fetchImpl = createFetchMock();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("issues a GET against the joined base URL and parses JSON", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(200, { value: 42 }));
    const client = new HttpClient({ baseUrl: "http://api.example", fetchImpl });

    const body = await client.get<{ value: number }>("/api/user/preferences");

    expect(body).toEqual({ value: 42 });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
    const [input, init] = fetchImpl.mock.calls[0] as [string, Parameters<typeof fetch>[1]];
    expect(input).toBe("http://api.example/api/user/preferences");
    expect(init.method).toBe("GET");
    expect(init.body).toBeUndefined();
  });

  it("attaches the bearer token from the provider when present", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(200, {}));
    const client = new HttpClient({
      baseUrl: "http://api.example",
      fetchImpl,
      getAuthToken: () => "tok-1",
    });

    await client.get("/x");

    const [, init] = fetchImpl.mock.calls[0] as [string, Parameters<typeof fetch>[1]];
    const headers = new Headers(init.headers);
    expect(headers.get("authorization")).toBe("Bearer tok-1");
  });

  it("omits the Authorization header when the provider returns undefined", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(200, {}));
    const client = new HttpClient({
      baseUrl: "http://api.example",
      fetchImpl,
      getAuthToken: () => undefined,
    });

    await client.get("/x");

    const [, init] = fetchImpl.mock.calls[0] as [string, Parameters<typeof fetch>[1]];
    const headers = new Headers(init.headers);
    expect(headers.get("authorization")).toBeNull();
  });
});

describe("HttpClient.put", () => {
  let fetchImpl: FetchMock;

  beforeEach(() => {
    fetchImpl = createFetchMock();
  });

  it("serializes the JSON body and sets Content-Type", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(200, { ok: true }));
    const client = new HttpClient({ baseUrl: "http://api.example", fetchImpl });

    await client.put("/api/user/preferences", { correctionInstructions: "vigilá X" });

    const [, init] = fetchImpl.mock.calls[0] as [string, Parameters<typeof fetch>[1]];
    expect(init.method).toBe("PUT");
    expect(init.body).toBe(JSON.stringify({ correctionInstructions: "vigilá X" }));
    const headers = new Headers(init.headers);
    expect(headers.get("content-type")).toBe("application/json");
  });

  it("serializes a null body as JSON null", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(200, {}));
    const client = new HttpClient({ baseUrl: "http://api.example", fetchImpl });

    await client.put("/api/user/preferences", { correctionInstructions: null });

    const [, init] = fetchImpl.mock.calls[0] as [string, Parameters<typeof fetch>[1]];
    expect(init.body).toBe(JSON.stringify({ correctionInstructions: null }));
  });
});

describe("HttpClient — error handling", () => {
  let fetchImpl: FetchMock;

  beforeEach(() => {
    fetchImpl = createFetchMock();
  });

  it("throws an HttpError carrying the parsed JSON body for non-2xx responses", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(401, { error: "unauthenticated" }));
    const client = new HttpClient({ baseUrl: "http://api.example", fetchImpl });

    const promise = client.get("/x");

    await expect(promise).rejects.toBeInstanceOf(HttpError);
    await expect(promise).rejects.toMatchObject({
      status: 401,
      body: { error: "unauthenticated" },
    });
  });

  it("preserves a textual body when the response is not JSON", async () => {
    fetchImpl.mockResolvedValueOnce(makeTextResponse(500, "boom"));
    const client = new HttpClient({ baseUrl: "http://api.example", fetchImpl });

    await expect(client.get("/x")).rejects.toMatchObject({
      status: 500,
      body: "boom",
    });
  });

  it("returns null for empty 204 responses", async () => {
    fetchImpl.mockResolvedValueOnce(makeEmptyResponse(204));
    const client = new HttpClient({ baseUrl: "http://api.example", fetchImpl });

    await expect(client.get("/x")).resolves.toBeNull();
  });

  it("propagates network errors as-is", async () => {
    const networkError = new TypeError("Failed to fetch");
    fetchImpl.mockRejectedValueOnce(networkError);
    const client = new HttpClient({ baseUrl: "http://api.example", fetchImpl });

    await expect(client.get("/x")).rejects.toBe(networkError);
  });
});

describe("HttpClient — base URL handling", () => {
  let fetchImpl: FetchMock;

  beforeEach(() => {
    fetchImpl = createFetchMock();
  });

  it("strips a trailing slash from the base URL before joining", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(200, {}));
    const client = new HttpClient({ baseUrl: "http://api.example/", fetchImpl });

    await client.get("/x");

    const [input] = fetchImpl.mock.calls[0] as [string];
    expect(input).toBe("http://api.example/x");
  });

  it("accepts a relative path without a leading slash", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(200, {}));
    const client = new HttpClient({ baseUrl: "http://api.example", fetchImpl });

    await client.get("api/user/preferences");

    const [input] = fetchImpl.mock.calls[0] as [string];
    expect(input).toBe("http://api.example/api/user/preferences");
  });
});
