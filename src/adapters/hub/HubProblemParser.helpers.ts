import { HttpError } from "../../infrastructure/http/HttpError";

/**
 * Reads the user-facing explanation of a failed hub request.
 *
 * The hub answers every refusal as `application/problem+json`
 * (`{ type, title, status, detail, code }`); `detail` carries the real cause.
 * Transport failures have no response, so their own message is the cause.
 */
export function readHubErrorDetail(error: unknown): string {
  if (error instanceof HttpError) {
    return (
      readProblemField(error.body, "detail") ??
      readProblemField(error.body, "title") ??
      error.message
    );
  }

  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message.trim();
  }

  return "error desconocido";
}

/** Returns the HTTP status of a failed hub request, or undefined when it never got one. */
export function readHubErrorStatus(error: unknown): number | undefined {
  return error instanceof HttpError ? error.status : undefined;
}

/** Reads a non-blank string field from a parsed problem body. */
function readProblemField(body: unknown, field: "detail" | "title"): string | undefined {
  if (typeof body !== "object" || body === null) {
    return undefined;
  }

  const value = (body as Record<string, unknown>)[field];
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
}
