import { describe, expect, it, vi } from "vitest";
import type {
  DocumentReviewState,
  ResolutionExecutionReport,
  SuggestionActionResult,
} from "../../../domain/types";
import type { ResolveSuggestionResultFactory } from "./ResolveSuggestionResultFactory";
import { SuggestionResolutionResolver } from "./SuggestionResolutionResolver";

function makeReviewState(
  overrides: Partial<DocumentReviewState> = {},
): DocumentReviewState {
  return {
    pendingStylisticArtifacts: 1,
    hasPendingStylisticArtifacts: true,
    trackChangesActive: true,
    ...overrides,
  };
}

function makeExecutionReport(
  overrides: Partial<ResolutionExecutionReport> = {},
): ResolutionExecutionReport {
  return {
    attempted: 2,
    completed: 2,
    remaining: 0,
    ...overrides,
  };
}

function makeResultFactoryDouble(
  buildResolutionResult: ReturnType<typeof vi.fn>,
  status: SuggestionActionResult["status"] = "accepted",
): ResolveSuggestionResultFactory {
  return {
    toResolutionStatus: () => status,
    buildResolutionResult,
  } as unknown as ResolveSuggestionResultFactory;
}

describe("SuggestionResolutionResolver", () => {
  it("returns null when execution still had unresolved tracked changes", () => {
    const buildResolutionResult = vi.fn();
    const resolver = new SuggestionResolutionResolver(
      "s-1",
      makeResultFactoryDouble(buildResolutionResult),
    );

    const result = resolver.reconcileLateFailure(
      makeReviewState(),
      makeExecutionReport({ completed: 1, remaining: 1, error: "ItemNotFound" }),
      "ItemNotFound",
    );

    expect(result).toBeNull();
    expect(buildResolutionResult).not.toHaveBeenCalled();
  });

  it("returns terminal semantic success with warnings when execution had already completed", () => {
    const pendingAfter = makeReviewState({
      pendingStylisticArtifacts: 0,
      hasPendingStylisticArtifacts: false,
    });
    const executionReport = makeExecutionReport();
    const expectedResult: SuggestionActionResult = {
      status: "accepted",
      trackedChangesAffected: 2,
      commentDeleted: false,
      pendingAfter,
      documentState: "ready-to-disable-track-changes",
      warnings: [],
      executionReport,
    };

    const buildResolutionResult = vi.fn(() => expectedResult);
    const resolver = new SuggestionResolutionResolver(
      "s-1",
      makeResultFactoryDouble(buildResolutionResult),
    );

    const result = resolver.reconcileLateFailure(
      pendingAfter,
      executionReport,
      "ItemNotFound",
      [
        {
          code: "telemetry-failed",
          phase: "execute",
          message: "telemetry sink offline",
        },
      ],
    );

    expect(result).toBe(expectedResult);
    expect(buildResolutionResult).toHaveBeenCalledWith(
      "accepted",
      2,
      false,
      pendingAfter,
      pendingAfter,
      undefined,
      [
        {
          code: "telemetry-failed",
          phase: "execute",
          message: "telemetry sink offline",
        },
        {
          code: "cleanup-failed",
          phase: "cleanup",
          message: "ItemNotFound",
        },
        {
          code: "reconciled-after-error",
          phase: "reconcile",
          message: "ItemNotFound",
        },
      ],
      executionReport,
    );
  });
});
