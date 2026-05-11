/** Stable, non-throwing diagnostic shape for unknown Office.js-ish errors. */
export type SerializedOfficeErrorDiagnostics = {
  message: string;
  name?: string;
  code?: string | number;
  debugInfo?: unknown;
  traceMessages?: unknown;
  stackPreview?: string[];
};
