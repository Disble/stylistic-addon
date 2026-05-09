import type { HttpErrorPayload } from "./HttpClient.types";

/**
 * Typed error raised by `HttpClient` for non-2xx responses.
 *
 * Adapters branch on `status` to translate transport failures into
 * domain-meaningful errors (e.g., 401 → unauthenticated, 400 → invalid request).
 * The original parsed body is preserved so callers can surface backend issues
 * such as validation `issues[]` payloads.
 */
export class HttpError extends Error implements HttpErrorPayload {
  public readonly status: number;
  public readonly statusText: string;
  public readonly body: unknown;

  constructor(payload: HttpErrorPayload) {
    super(`HTTP ${payload.status} ${payload.statusText}`);
    this.name = "HttpError";
    this.status = payload.status;
    this.statusText = payload.statusText;
    this.body = payload.body;
  }
}
