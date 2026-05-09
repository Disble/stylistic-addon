/**
 * HTTP infrastructure contracts.
 *
 * The HttpClient layer is intentionally framework-free: adapters consume it via
 * a narrow contract, not via raw `fetch`. Swapping the implementation for axios
 * or adding interceptors should require only changes inside `HttpClient.ts`.
 */

/** Supplies the current Better Auth bearer token for authenticated HTTP calls. */
export type HttpAuthTokenProvider = () => string | undefined;

/** Supported HTTP methods used by the project's REST adapters. */
export type HttpMethod = "GET" | "POST" | "PUT" | "PATCH" | "DELETE";

/** Optional per-request configuration accepted by `HttpClient`. */
export type HttpRequestOptions = Readonly<{
  /** Optional JSON body. The client serializes it and sets Content-Type. */
  body?: unknown;
  /** Extra headers merged on top of the default Authorization/Accept set. */
  headers?: Readonly<Record<string, string>>;
  /** AbortSignal forwarded to `fetch` so callers can cancel in-flight requests. */
  signal?: AbortSignal;
}>;

/** Constructor dependencies for `HttpClient`. */
export type HttpClientDeps = Readonly<{
  /** Base URL prepended to every relative path (no trailing slash). */
  baseUrl: string;
  /** Returns the latest bearer token snapshot, or undefined when anonymous. */
  getAuthToken?: HttpAuthTokenProvider;
  /** Optional `fetch` override used in tests. Defaults to `globalThis.fetch`. */
  fetchImpl?: typeof fetch;
}>;

/** Structured failure raised by `HttpClient` for non-2xx responses. */
export type HttpErrorPayload = Readonly<{
  status: number;
  statusText: string;
  body: unknown;
}>;
