/**
 * Reads one property from an unknown record without throwing.
 */
export function readUnknownRecordProperty(error: unknown, propertyName: string): unknown {
  if (typeof error !== "object" || error === null) {
    return undefined;
  }

  try {
    return (error as Record<string, unknown>)[propertyName];
  } catch {
    return undefined;
  }
}

/**
 * Converts unknown runtime failures into stable diagnostics without opaque object stringification.
 */
export function stringifyUnknownError(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  if (typeof error === "string") {
    return error;
  }

  const primitive = stringifyUnknownPrimitive(error);
  if (primitive) {
    return primitive;
  }

  const message = readUnknownRecordProperty(error, "message");
  if (typeof message === "string" && message.trim().length > 0) {
    return message;
  }

  const code = readUnknownRecordProperty(error, "code");
  if (typeof code === "string" || typeof code === "number") {
    return String(code);
  }

  if (typeof error === "object" && error !== null) {
    try {
      const serialized = JSON.stringify(error);
      if (serialized && serialized !== "{}") {
        return serialized;
      }
    } catch {
      return "Unknown object error";
    }

    return "Unknown object error";
  }

  return "Unknown error";
}

/**
 * Converts primitive thrown values to stable text.
 */
function stringifyUnknownPrimitive(error: unknown): string | null {
  if (typeof error === "number" || typeof error === "boolean" || typeof error === "bigint") {
    return String(error);
  }

  if (typeof error === "symbol") {
    return error.description ? `Symbol(${error.description})` : "Symbol()";
  }

  return null;
}
