import { describe, expect, it, vi } from "vitest";
import type { Suggestion } from "../../../../domain/suggestion/Suggestion.types";
import type { DocumentReviewStateInspector } from "../DocumentReviewStateInspector";
import type { ResolutionObservation } from "../ResolutionContext";
import type { ResolutionErrorSerializer } from "../ResolutionErrorParser";
import type { ResolutionObservabilityReporter } from "../ResolutionObservabilityAdapter";
import { ResolveSuggestionTrackChangeOrchestrator } from "../ResolveSuggestionTrackChangeOrchestrator";
import type { ResolveSuggestionResultFactory } from "../ResolveSuggestionResultFactory";
import type { SuggestionLocator } from "../SuggestionLocator";
import type { SuggestionResolutionCleanup } from "../SuggestionResolutionCleanup";
import type { SuggestionResolutionObserver } from "../SuggestionResolutionObserver";

const pendingBefore = {
  pendingStylisticArtifacts: 1,
  hasPendingStylisticArtifacts: true,
  trackChangesActive: true,
};

const pendingAfter = {
  pendingStylisticArtifacts: 0,
  hasPendingStylisticArtifacts: false,
  trackChangesActive: true,
};

/** Builds the minimum valid track-change suggestion contract for orchestration tests. */
function makeTrackChangeSuggestion(): Suggestion {
  return {
    id: "s-cleanup",
    context: "El texto original.",
    anchor: "original",
    suggestedText: "nuevo",
    justification: "Evita ambigüedad.",
    category: "claridad",
    severity: "medium",
    type: "track-change",
  };
}

function makeDeleteOnlySuggestion(): Suggestion {
  return {
    ...makeTrackChangeSuggestion(),
    id: "s-delete",
    anchor: " a pesar de eso",
    context: "No obstante, siguió sosteniéndola del brazo a pesar de eso.",
    suggestedText: "",
  };
}

function makeFormattingSuggestion(): Suggestion {
  return {
    ...makeTrackChangeSuggestion(),
    id: "s-format",
    anchor: "post mortem",
    context: "Ese era el inicio del post mortem reportado por PRIME.",
    suggestedText: "*post mortem*",
  };
}

/** Converts unknown thrown values to stable test diagnostics. */
function stringifyUnknownError(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  if (typeof error === "string") {
    return error;
  }

  if (typeof error === "number" || typeof error === "boolean" || typeof error === "bigint") {
    return String(error);
  }

  return "Unknown error";
}

describe("ResolveSuggestionTrackChangeOrchestrator metadata cleanup", () => {
  it("cleans resolved metadata before final document state inspection", async () => {
    const order: string[] = [];
    const context = { sync: vi.fn(async () => undefined) } as unknown as Word.RequestContext;
    const selectedCc = {} as Word.ContentControl;
    const selectedComment = {
      comment: {} as Word.Comment,
      range: {} as Word.Range,
    };
    const trackedChangesCollection = {
      items: [{} as Word.TrackedChange],
      load: vi.fn(),
      acceptAll: vi.fn(() => order.push("execute")),
      rejectAll: vi.fn(),
    };
    const observation: ResolutionObservation = {
      selectedCc,
      selectedComment,
      trackedChanges: [{} as Word.TrackedChange],
      trackedChangesCollection,
      observationStatus: "confirmed-pending",
    };

    const locator = {
      locateResolutionArtifacts: vi.fn(async () => ({
        candidates: [selectedCc],
        selectedCc,
        locateStatus: "confirmed-pending",
      })),
    } as unknown as SuggestionLocator;
    const observer = {
      observeResolutionCandidates: vi.fn(async () => observation),
    } as unknown as SuggestionResolutionObserver;
    const cleanup = {
      deleteLocatedStylisticCommentAfterResolution: vi.fn(async () => {
        order.push("cleanup-comment");
        return true;
      }),
      deleteResolvedTrackChangeMetadata: vi.fn(async () => {
        order.push("cleanup-metadata");
        return {
          deletedContentControls: [
            "stylistic-operational-wrapper:s-cleanup",
            "stylistic:track-change:s-cleanup",
          ],
          failedContentControls: [],
        };
      }),
    } as unknown as SuggestionResolutionCleanup;
    const stateInspector = {
      inspect: vi.fn(async () => pendingBefore),
      inspectAfterResolution: vi.fn(async () => {
        order.push("inspect-after");
        return pendingAfter;
      }),
    } as unknown as DocumentReviewStateInspector;
    const resultFactory = {
      toResolutionStatus: vi.fn(() => "accepted"),
      buildErrorResult: vi.fn(),
      buildObservationFailureResult: vi.fn(),
    } as unknown as ResolveSuggestionResultFactory;
    const reporter = {
      emitPhase: vi.fn(async () => undefined),
      mergeMetadata: vi.fn((base, extra) => ({ ...base, ...extra })),
    } as unknown as ResolutionObservabilityReporter;
    const errorSerializer = {
      serialize: vi.fn((error: unknown) => ({
        message: stringifyUnknownError(error),
      })),
    } as unknown as ResolutionErrorSerializer;
    const orchestrator = new ResolveSuggestionTrackChangeOrchestrator(
      makeTrackChangeSuggestion(),
      "accept",
      locator,
      cleanup,
      observer,
      resultFactory,
      stateInspector,
      reporter,
      errorSerializer
    );

    const result = await orchestrator.execute(context);

    expect(order).toEqual(["execute", "cleanup-comment", "cleanup-metadata", "inspect-after"]);
    expect(cleanup.deleteResolvedTrackChangeMetadata).toHaveBeenCalledWith(context);
    expect(result.status).toBe("accepted");
    expect(result.pendingAfter).toBe(pendingAfter);
    expect(reporter.emitPhase).toHaveBeenCalledWith(
      "cleanup-metadata",
      "succeeded",
      expect.objectContaining({ deletedContentControlCount: 2 })
    );
  });

  it("accepts delete-only track-change contracts with empty suggestedText", async () => {
    const context = { sync: vi.fn(async () => undefined) } as unknown as Word.RequestContext;
    const selectedCc = {} as Word.ContentControl;
    const observation: ResolutionObservation = {
      selectedCc,
      selectedComment: null,
      trackedChanges: [{} as Word.TrackedChange],
      trackedChangesCollection: {
        items: [{} as Word.TrackedChange],
        load: vi.fn(),
        acceptAll: vi.fn(),
        rejectAll: vi.fn(),
      },
      observationStatus: "confirmed-pending",
    };
    const locator = {
      locateResolutionArtifacts: vi.fn(async () => ({
        candidates: [selectedCc],
        selectedCc,
        locateStatus: "confirmed-pending",
      })),
    } as unknown as SuggestionLocator;
    const observer = {
      observeResolutionCandidates: vi.fn(async () => observation),
    } as unknown as SuggestionResolutionObserver;
    const cleanup = {
      deleteLocatedStylisticCommentAfterResolution: vi.fn(async () => false),
      deleteResolvedTrackChangeMetadata: vi.fn(async () => ({
        deletedContentControls: [],
        failedContentControls: [],
      })),
    } as unknown as SuggestionResolutionCleanup;
    const stateInspector = {
      inspect: vi.fn(async () => pendingBefore),
      inspectAfterResolution: vi.fn(async () => pendingAfter),
    } as unknown as DocumentReviewStateInspector;
    const resultFactory = {
      toResolutionStatus: vi.fn(() => "accepted"),
      buildErrorResult: vi.fn(),
      buildObservationFailureResult: vi.fn(),
    } as unknown as ResolveSuggestionResultFactory;
    const reporter = {
      emitPhase: vi.fn(async () => undefined),
      mergeMetadata: vi.fn((base, extra) => ({ ...base, ...extra })),
    } as unknown as ResolutionObservabilityReporter;
    const errorSerializer = {
      serialize: vi.fn((error: unknown) => ({
        message: stringifyUnknownError(error),
      })),
    } as unknown as ResolutionErrorSerializer;

    const orchestrator = new ResolveSuggestionTrackChangeOrchestrator(
      makeDeleteOnlySuggestion(),
      "accept",
      locator,
      cleanup,
      observer,
      resultFactory,
      stateInspector,
      reporter,
      errorSerializer
    );

    const result = await orchestrator.execute(context);

    expect(result.status).toBe("accepted");
    expect(resultFactory.buildErrorResult).not.toHaveBeenCalled();
  });

  it("accepts formatting track-change contracts encoded as markdown", async () => {
    const context = { sync: vi.fn(async () => undefined) } as unknown as Word.RequestContext;
    const selectedCc = {} as Word.ContentControl;
    const observation: ResolutionObservation = {
      selectedCc,
      selectedComment: null,
      trackedChanges: [{} as Word.TrackedChange],
      trackedChangesCollection: {
        items: [{} as Word.TrackedChange],
        load: vi.fn(),
        acceptAll: vi.fn(),
        rejectAll: vi.fn(),
      },
      observationStatus: "confirmed-pending",
    };
    const locator = {
      locateResolutionArtifacts: vi.fn(async () => ({
        candidates: [selectedCc],
        selectedCc,
        locateStatus: "confirmed-pending",
      })),
    } as unknown as SuggestionLocator;
    const observer = {
      observeResolutionCandidates: vi.fn(async () => observation),
    } as unknown as SuggestionResolutionObserver;
    const cleanup = {
      deleteLocatedStylisticCommentAfterResolution: vi.fn(async () => false),
      deleteResolvedTrackChangeMetadata: vi.fn(async () => ({
        deletedContentControls: [],
        failedContentControls: [],
      })),
    } as unknown as SuggestionResolutionCleanup;
    const stateInspector = {
      inspect: vi.fn(async () => pendingBefore),
      inspectAfterResolution: vi.fn(async () => pendingAfter),
    } as unknown as DocumentReviewStateInspector;
    const resultFactory = {
      toResolutionStatus: vi.fn(() => "accepted"),
      buildErrorResult: vi.fn(),
      buildObservationFailureResult: vi.fn(),
    } as unknown as ResolveSuggestionResultFactory;
    const reporter = {
      emitPhase: vi.fn(async () => undefined),
      mergeMetadata: vi.fn((base, extra) => ({ ...base, ...extra })),
    } as unknown as ResolutionObservabilityReporter;
    const errorSerializer = {
      serialize: vi.fn((error: unknown) => ({
        message: stringifyUnknownError(error),
      })),
    } as unknown as ResolutionErrorSerializer;

    const orchestrator = new ResolveSuggestionTrackChangeOrchestrator(
      makeFormattingSuggestion(),
      "accept",
      locator,
      cleanup,
      observer,
      resultFactory,
      stateInspector,
      reporter,
      errorSerializer
    );

    const result = await orchestrator.execute(context);

    expect(result.status).toBe("accepted");
    expect(resultFactory.buildErrorResult).not.toHaveBeenCalled();
  });
});
