import {
  readUnknownRecordProperty,
  stringifyUnknownError,
} from "../../../infrastructure/errors/UnknownError.helpers";
import type { SerializedOfficeErrorDiagnostics } from "./ResolutionErrorParser.types";

/** Converts unknown runtime errors into safe diagnostic objects for logs and observability. */
export class ResolutionErrorSerializer {
  /** Converts one unknown Office.js-ish error into a stable message without stringifying opaque objects. */
  stringify(error: unknown): string {
    return stringifyUnknownError(error);
  }

  /** Builds one plain Office.js-ish error diagnostic object for console output. */
  serialize(error: unknown): SerializedOfficeErrorDiagnostics {
    const fallbackMessage = this.stringify(error);
    const messageValue = readUnknownRecordProperty(error, "message");
    const nameValue = readUnknownRecordProperty(error, "name");
    const codeValue = readUnknownRecordProperty(error, "code");
    const debugInfo = readUnknownRecordProperty(error, "debugInfo");
    const traceMessages = readUnknownRecordProperty(error, "traceMessages");
    const stackValue = readUnknownRecordProperty(error, "stack");
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
