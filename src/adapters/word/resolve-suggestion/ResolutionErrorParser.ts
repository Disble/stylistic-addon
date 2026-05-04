/** Stable, non-throwing diagnostic shape for unknown Office.js-ish errors. */
export type SerializedOfficeErrorDiagnostics = {
  message: string;
  name?: string;
  code?: string | number;
  debugInfo?: unknown;
  traceMessages?: unknown;
  stackPreview?: string[];
};

/** Converts unknown runtime errors into safe diagnostic objects for logs and observability. */
export class ResolutionErrorSerializer {
  /** Reads one unknown error property defensively so diagnostic logging never throws. */
  private readUnknownErrorProperty(error: unknown, propertyName: string): unknown {
    if (typeof error !== "object" || error === null) {
      return undefined;
    }

    try {
      return (error as Record<string, unknown>)[propertyName];
    } catch {
      return undefined;
    }
  }

  /** Converts one unknown Office.js-ish error into a stable message without stringifying opaque objects. */
  stringify(error: unknown): string {
    if (error instanceof Error) {
      return error.message;
    }

    if (typeof error === "string") {
      return error;
    }

    if (typeof error === "number" || typeof error === "boolean" || typeof error === "bigint") {
      return String(error);
    }

    const messageValue = this.readUnknownErrorProperty(error, "message");
    if (typeof messageValue === "string" && messageValue.trim().length > 0) {
      return messageValue;
    }

    return "Unknown error";
  }

  /** Builds one plain Office.js-ish error diagnostic object for console output. */
  serialize(error: unknown): SerializedOfficeErrorDiagnostics {
    const fallbackMessage = this.stringify(error);
    const messageValue = this.readUnknownErrorProperty(error, "message");
    const nameValue = this.readUnknownErrorProperty(error, "name");
    const codeValue = this.readUnknownErrorProperty(error, "code");
    const debugInfo = this.readUnknownErrorProperty(error, "debugInfo");
    const traceMessages = this.readUnknownErrorProperty(error, "traceMessages");
    const stackValue = this.readUnknownErrorProperty(error, "stack");
    const stackPreview =
      typeof stackValue === "string"
        ? stackValue
            .split(/\r?\n/u)
            .map((line) => line.trim())
            .filter((line) => line.length > 0)
            .slice(0, 5)
        : undefined;

    const serializedError: SerializedOfficeErrorDiagnostics = {
      message:
        typeof messageValue === "string" && messageValue.length > 0
          ? messageValue
          : fallbackMessage,
    };

    if (typeof nameValue === "string" && nameValue.length > 0) {
      serializedError.name = nameValue;
    }

    if (typeof codeValue === "string" || typeof codeValue === "number") {
      serializedError.code = codeValue;
    }

    if (debugInfo !== undefined) {
      serializedError.debugInfo = debugInfo;
    }

    if (traceMessages !== undefined) {
      serializedError.traceMessages = traceMessages;
    }

    if (stackPreview && stackPreview.length > 0) {
      serializedError.stackPreview = stackPreview;
    }

    return serializedError;
  }
}
