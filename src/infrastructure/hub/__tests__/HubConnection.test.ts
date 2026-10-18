import { describe, expect, it } from "vitest";
import { buildHubApiBaseUrl, readHubToken } from "../HubConnection.helpers";

/** Builds a minimal document whose only meta tag is the hub token with the given content. */
function documentWithHubMeta(content: string | null): Pick<Document, "querySelector"> {
  return {
    querySelector: (selector: string) =>
      selector === 'meta[name="stylistic-hub-token"]'
        ? ({ getAttribute: (name: string) => (name === "content" ? content : null) } as Element)
        : null,
  } as Pick<Document, "querySelector">;
}

describe("readHubToken", () => {
  it("returns the hub token when the meta tag has content", () => {
    expect(readHubToken(documentWithHubMeta(` ${"a".repeat(64)} `))).toBe("a".repeat(64));
  });

  it("returns undefined when the meta tag is empty", () => {
    expect(readHubToken(documentWithHubMeta("  "))).toBeUndefined();
  });

  it("returns undefined when the meta tag is missing", () => {
    expect(readHubToken({ querySelector: () => null })).toBeUndefined();
  });

  it("returns undefined when there is no document", () => {
    expect(readHubToken(undefined)).toBeUndefined();
  });
});

describe("buildHubApiBaseUrl", () => {
  it("appends the API v1 path to the task pane origin", () => {
    expect(buildHubApiBaseUrl("https://localhost:3000")).toBe("https://localhost:3000/api/v1");
  });

  it("does not double a trailing slash", () => {
    expect(buildHubApiBaseUrl("https://localhost:3000/")).toBe("https://localhost:3000/api/v1");
  });
});
