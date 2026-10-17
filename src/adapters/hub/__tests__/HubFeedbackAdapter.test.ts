/* global RequestInit */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { HttpClient } from "../../../infrastructure/http/HttpClient";
import { HubFeedbackAdapter } from "../HubFeedbackAdapter";

type FetchMock = ReturnType<typeof vi.fn> & typeof fetch;

const BASE_URL = "https://localhost:3000/api/v1";

describe("HubFeedbackAdapter", () => {
  let fetchImpl: FetchMock;
  let adapter: HubFeedbackAdapter;

  beforeEach(() => {
    fetchImpl = vi.fn() as unknown as FetchMock;
    adapter = new HubFeedbackAdapter(
      new HttpClient({ baseUrl: BASE_URL, fetchImpl, getAuthToken: () => "hub-token" })
    );
    vi.spyOn(console, "log").mockImplementation(() => {});
    vi.spyOn(console, "error").mockImplementation(() => {});
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("posts the suggestion id, action and comment to /feedback", async () => {
    fetchImpl.mockResolvedValueOnce(
      new Response(JSON.stringify({ id: "f1", status: "pending" }), {
        status: 202,
        headers: { "content-type": "application/json" },
      })
    );

    await adapter.sendFeedback({
      suggestionId: "0123456789abcdef",
      action: "reject",
      comment: "No aplica en este registro.",
    });

    const [url, init] = fetchImpl.mock.calls[0] as [string, RequestInit];
    expect(url).toBe(`${BASE_URL}/feedback`);
    expect(init.method).toBe("POST");
    expect(new Headers(init.headers).get("Authorization")).toBe("Bearer hub-token");
    expect(JSON.parse(init.body as string)).toEqual({
      suggestionId: "0123456789abcdef",
      action: "reject",
      comment: "No aplica en este registro.",
    });
  });

  it("swallows hub refusals so feedback never surfaces to the user", async () => {
    fetchImpl.mockResolvedValueOnce(
      new Response(
        JSON.stringify({ status: 404, code: "not_found", detail: "No such suggestion" }),
        {
          status: 404,
          headers: { "content-type": "application/problem+json" },
        }
      )
    );

    await expect(
      adapter.sendFeedback({ suggestionId: "0123456789abcdef", action: "accept" })
    ).resolves.toBeUndefined();
  });

  it("swallows transport failures", async () => {
    fetchImpl.mockRejectedValueOnce(new TypeError("Failed to fetch"));

    await expect(
      adapter.sendFeedback({ suggestionId: "0123456789abcdef", action: "accept" })
    ).resolves.toBeUndefined();
  });
});
