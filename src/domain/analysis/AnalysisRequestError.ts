/**
 * Failure thrown by adapters implementing `IAnalysisPort`.
 *
 * `message` carries the backend's own explanation so the taskpane can show the
 * real cause. `status` is the HTTP status when the backend answered, and is
 * absent for transport failures (no response).
 */
export class AnalysisRequestError extends Error {
  public readonly status?: number;

  constructor(message: string, status?: number) {
    super(message);
    this.name = "AnalysisRequestError";
    this.status = status;
  }
}
