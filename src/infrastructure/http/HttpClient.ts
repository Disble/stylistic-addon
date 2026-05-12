/* global fetch, Headers, Response, RequestInit, console */

import { HttpError } from "./HttpError";
import type {
  HttpClientDeps,
  HttpAuthTokenProvider,
  HttpMethod,
  HttpRequestOptions,
} from "./HttpClient.types";

/**
 * HTTP client — thin layer on top of `fetch` used by REST adapters.
 *
 * Responsibilities:
 * - Join the base URL and a relative path without surprising slashes.
 * - Inject the Better Auth bearer token through the configured provider.
 * - Serialize JSON request bodies and parse JSON or text responses.
 * - Translate non-2xx responses into a typed {@link HttpError}.
 *
 * Adapters MUST depend on this class instead of calling `fetch` directly so
 * cross-cutting concerns (auth, retries, observability) can grow in one place.
 */
export class HttpClient {
  private readonly baseUrl: string;
  private readonly getAuthToken: HttpAuthTokenProvider;
  private readonly fetchImpl: typeof fetch;

  constructor(deps: HttpClientDeps) {
    this.baseUrl = this.normalizeBaseUrl(deps.baseUrl);
    this.getAuthToken = deps.getAuthToken ?? (() => undefined);
    this.fetchImpl = deps.fetchImpl ?? globalThis.fetch.bind(globalThis);
  }

  /** Removes trailing slashes without regex backtracking so base URL joins stay deterministic. */
  private normalizeBaseUrl(baseUrl: string): string {
    let endIndex = baseUrl.length;

    while (endIndex > 0 && baseUrl.codePointAt(endIndex - 1) === 47) {
      endIndex -= 1;
    }

    return baseUrl.slice(0, endIndex);
  }

  /** Formats unknown thrown values so logs never stringify opaque objects accidentally. */
  private describeThrownError(error: unknown): string {
    if (error instanceof Error) {
      return error.message;
    }

    if (typeof error === "string") {
      return error;
    }

    if (typeof error === "number") {
      return `${error}`;
    }

    if (typeof error === "boolean") {
      return error ? "true" : "false";
    }

    if (typeof error === "bigint") {
      return error.toString();
    }

    if (error && typeof error === "object") {
      const message = (error as { message?: unknown }).message;
      if (typeof message === "string" && message.length > 0) {
        return message;
      }
      return "Unknown object error";
    }

    return "Unknown error";
  }

  async get<TResponse>(path: string, options?: HttpRequestOptions): Promise<TResponse> {
    return this.request<TResponse>("GET", path, options);
  }

  async put<TResponse>(
    path: string,
    body: unknown,
    options?: HttpRequestOptions
  ): Promise<TResponse> {
    return this.request<TResponse>("PUT", path, { ...options, body });
  }

  async post<TResponse>(
    path: string,
    body: unknown,
    options?: HttpRequestOptions
  ): Promise<TResponse> {
    return this.request<TResponse>("POST", path, { ...options, body });
  }

  private async request<TResponse>(
    method: HttpMethod,
    path: string,
    options?: HttpRequestOptions
  ): Promise<TResponse> {
    const url = this.buildUrl(path);
    const headers = this.buildHeaders(options?.headers, options?.body !== undefined);

    const init: RequestInit = {
      method,
      headers,
      signal: options?.signal,
    };

    if (options?.body !== undefined) {
      init.body = JSON.stringify(options.body);
    }

    console.log(`🌐 [HttpClient] ${method} ${url}`);
    try {
      const response = await this.fetchImpl(url, init);
      console.log(`🌐 [HttpClient] ${method} ${url} → ${response.status} ${response.statusText}`);
      return await this.handleResponse<TResponse>(response);
    } catch (error) {
      if (error instanceof HttpError) {
        console.error(
          `🔴 [HttpClient] ${method} ${url} failed with HTTP ${error.status}:`,
          error.body
        );
      } else {
        console.error(
          `🔴 [HttpClient] ${method} ${url} threw before reaching server: ${this.describeThrownError(error)}`
        );
      }
      throw error;
    }
  }

  private buildUrl(path: string): string {
    const normalizedPath = path.startsWith("/") ? path : `/${path}`;
    return `${this.baseUrl}${normalizedPath}`;
  }

  private buildHeaders(
    extra: Readonly<Record<string, string>> | undefined,
    hasBody: boolean
  ): Headers {
    const headers = new Headers();
    headers.set("Accept", "application/json");
    if (hasBody) {
      headers.set("Content-Type", "application/json");
    }

    const token = this.getAuthToken();
    if (token) {
      headers.set("Authorization", `Bearer ${token}`);
    }

    if (extra) {
      for (const [key, value] of Object.entries(extra)) {
        headers.set(key, value);
      }
    }
    return headers;
  }

  private async handleResponse<TResponse>(response: Response): Promise<TResponse> {
    const body = await this.parseResponseBody(response);

    if (!response.ok) {
      throw new HttpError({
        status: response.status,
        statusText: response.statusText,
        body,
      });
    }

    return body as TResponse;
  }

  private async parseResponseBody(response: Response): Promise<unknown> {
    if (response.status === 204) {
      return null;
    }
    const contentType = response.headers.get("content-type") ?? "";
    if (contentType.includes("application/json")) {
      return response.json();
    }
    const text = await response.text();
    return text.length === 0 ? null : text;
  }
}
