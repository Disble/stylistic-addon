import { HUB_API_PATH, HUB_TOKEN_META_NAME } from "./HubConnection.constants";

/**
 * Returns the hub bearer token from the loaded task pane document.
 *
 * The hub serves the task pane and injects its token as
 * `<meta name="stylistic-hub-token" content="...">`. Returns undefined when the
 * tag is missing or empty, which means the task pane was not served by a hub.
 */
export function readHubToken(doc: Pick<Document, "querySelector"> | undefined): string | undefined {
  const content = doc
    ?.querySelector(`meta[name="${HUB_TOKEN_META_NAME}"]`)
    ?.getAttribute("content")
    ?.trim();
  return content || undefined;
}

/** Builds the hub API base URL from the origin that served the task pane. */
export function buildHubApiBaseUrl(origin: string): string {
  return `${origin.replace(/\/+$/, "")}${HUB_API_PATH}`;
}
