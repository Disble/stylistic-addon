/** Converts unknown error values into a stable string. */
export function stringifyUnknownError(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  if (typeof error === "string") {
    return error;
  }

  try {
    return JSON.stringify(error);
  } catch {
    return Object.prototype.toString.call(error);
  }
}
