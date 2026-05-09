import { beforeEach, describe, expect, it, vi } from "vitest";
import { HttpClient } from "../../../infrastructure/http/HttpClient";
import { HttpError } from "../../../infrastructure/http/HttpError";
import { UserCorrectionPreferencesError } from "../../../domain/user-preferences/UserCorrectionPreferencesError";
import { BackendUserCorrectionPreferencesAdapter } from "../BackendUserCorrectionPreferencesAdapter";
import { USER_PREFERENCES_PATH } from "../BackendUserCorrectionPreferencesAdapter.constants";

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

function makeAdapter(fetchImpl: FetchMock): BackendUserCorrectionPreferencesAdapter {
  const httpClient = new HttpClient({ baseUrl: "http://api.example", fetchImpl });
  return new BackendUserCorrectionPreferencesAdapter(httpClient);
}

describe("BackendUserCorrectionPreferencesAdapter.load", () => {
  let fetchImpl: FetchMock;

  beforeEach(() => {
    fetchImpl = createFetchMock();
  });

  it("issues a GET against the user preferences path", async () => {
    fetchImpl.mockResolvedValueOnce(
      makeJsonResponse(200, {
        correctionInstructions: null,
        correctionInstructionsMaxLength: 4000,
      })
    );

    const adapter = makeAdapter(fetchImpl);
    await adapter.load();

    const [url, init] = fetchImpl.mock.calls[0] as [string, Parameters<typeof fetch>[1]];
    expect(url).toBe(`http://api.example${USER_PREFERENCES_PATH}`);
    expect(init.method).toBe("GET");
  });

  it("returns the parsed preferences when present", async () => {
    fetchImpl.mockResolvedValueOnce(
      makeJsonResponse(200, {
        correctionInstructions: "Vigilá muletillas.",
        correctionInstructionsMaxLength: 4000,
      })
    );

    const adapter = makeAdapter(fetchImpl);

    await expect(adapter.load()).resolves.toEqual({
      correctionInstructions: "Vigilá muletillas.",
      correctionInstructionsMaxLength: 4000,
    });
  });

  it("translates 401 into an unauthenticated domain error", async () => {
    fetchImpl.mockImplementation(async () => makeJsonResponse(401, { error: "unauthenticated" }));

    const adapter = makeAdapter(fetchImpl);

    await expect(adapter.load()).rejects.toBeInstanceOf(UserCorrectionPreferencesError);
    await expect(adapter.load()).rejects.toMatchObject({ reason: "unauthenticated" });
  });

  it("translates network errors into a network domain error", async () => {
    fetchImpl.mockRejectedValueOnce(new TypeError("Failed to fetch"));

    const adapter = makeAdapter(fetchImpl);

    await expect(adapter.load()).rejects.toMatchObject({ reason: "network" });
  });
});

describe("BackendUserCorrectionPreferencesAdapter.save", () => {
  let fetchImpl: FetchMock;

  beforeEach(() => {
    fetchImpl = createFetchMock();
  });

  it("issues a PUT carrying the provided correctionInstructions string", async () => {
    fetchImpl.mockResolvedValueOnce(
      makeJsonResponse(200, {
        correctionInstructions: "Vigilá subordinadas largas.",
        correctionInstructionsMaxLength: 4000,
      })
    );

    const adapter = makeAdapter(fetchImpl);
    await adapter.save("Vigilá subordinadas largas.");

    const [url, init] = fetchImpl.mock.calls[0] as [string, Parameters<typeof fetch>[1]];
    expect(url).toBe(`http://api.example${USER_PREFERENCES_PATH}`);
    expect(init.method).toBe("PUT");
    expect(init.body).toBe(
      JSON.stringify({ correctionInstructions: "Vigilá subordinadas largas." })
    );
  });

  it("sends null when clearing the stored preferences", async () => {
    fetchImpl.mockResolvedValueOnce(
      makeJsonResponse(200, {
        correctionInstructions: null,
        correctionInstructionsMaxLength: 4000,
      })
    );

    const adapter = makeAdapter(fetchImpl);
    await adapter.save(null);

    const [, init] = fetchImpl.mock.calls[0] as [string, Parameters<typeof fetch>[1]];
    expect(init.body).toBe(JSON.stringify({ correctionInstructions: null }));
  });

  it("returns the backend's authoritative response (post-trim, post-collapse)", async () => {
    fetchImpl.mockResolvedValueOnce(
      makeJsonResponse(200, {
        correctionInstructions: null,
        correctionInstructionsMaxLength: 4000,
      })
    );

    const adapter = makeAdapter(fetchImpl);

    await expect(adapter.save("   ")).resolves.toEqual({
      correctionInstructions: null,
      correctionInstructionsMaxLength: 4000,
    });
  });

  it("translates 400 invalid_user_preferences_request into invalid-request", async () => {
    fetchImpl.mockResolvedValueOnce(
      makeJsonResponse(400, {
        error: "invalid_user_preferences_request",
        issues: [{ code: "too_big" }],
      })
    );

    const adapter = makeAdapter(fetchImpl);

    await expect(adapter.save("x".repeat(4001))).rejects.toMatchObject({
      reason: "invalid-request",
    });
  });

  it("translates 401 into unauthenticated", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(401, { error: "unauthenticated" }));

    const adapter = makeAdapter(fetchImpl);

    await expect(adapter.save("anything")).rejects.toMatchObject({
      reason: "unauthenticated",
    });
  });

  it("translates other HTTP errors into unknown", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(500, { error: "boom" }));

    const adapter = makeAdapter(fetchImpl);

    await expect(adapter.save("anything")).rejects.toMatchObject({
      reason: "unknown",
    });
  });

  it("preserves UserCorrectionPreferencesError instances when raised by the http client", async () => {
    fetchImpl.mockResolvedValueOnce(makeJsonResponse(401, { error: "unauthenticated" }));

    const adapter = makeAdapter(fetchImpl);

    await expect(adapter.save("anything")).rejects.toBeInstanceOf(UserCorrectionPreferencesError);
  });
});

describe("BackendUserCorrectionPreferencesAdapter — error normalization helper", () => {
  it("treats any non-HttpError as an unknown reason", async () => {
    const fetchImpl = vi.fn().mockRejectedValueOnce(new Error("some unexpected"));
    const adapter = makeAdapter(fetchImpl);

    await expect(adapter.load()).rejects.toMatchObject({ reason: "unknown" });
  });

  it("treats TypeError (network) distinctly from generic errors", async () => {
    const fetchImpl = vi.fn().mockRejectedValueOnce(new TypeError("offline"));
    const adapter = makeAdapter(fetchImpl);

    const result = adapter.load();
    await expect(result).rejects.toBeInstanceOf(UserCorrectionPreferencesError);
    await expect(result).rejects.toMatchObject({ reason: "network" });
  });

  it("does not double-wrap an already typed error from a prior layer", async () => {
    const httpClient = {
      get: vi
        .fn()
        .mockRejectedValue(new UserCorrectionPreferencesError("invalid-request", "pre-existing")),
      put: vi.fn(),
    } as unknown as HttpClient;

    const adapter = new BackendUserCorrectionPreferencesAdapter(httpClient);

    const promise = adapter.load();
    await expect(promise).rejects.toMatchObject({
      reason: "invalid-request",
      message: "pre-existing",
    });
  });

  it("wraps HttpError instances thrown directly by HttpClient", async () => {
    const httpError = new HttpError({
      status: 401,
      statusText: "Unauthorized",
      body: { error: "unauthenticated" },
    });
    const httpClient = {
      get: vi.fn().mockRejectedValue(httpError),
      put: vi.fn(),
    } as unknown as HttpClient;

    const adapter = new BackendUserCorrectionPreferencesAdapter(httpClient);

    await expect(adapter.load()).rejects.toMatchObject({
      reason: "unauthenticated",
    });
  });
});
