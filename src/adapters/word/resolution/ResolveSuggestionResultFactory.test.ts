import type { DocumentReviewUiState } from "../../../domain/review/DocumentReviewStateMachine";
import type { DocumentReviewState } from "../../../domain/types";
import { describe, expect, it, vi } from "vitest";
import { ResolveSuggestionResultFactory } from "./ResolveSuggestionResultFactory";

describe("ResolveSuggestionResultFactory", () => {
  it("builds an identity-lost result with the expected retry-safe error message", async () => {
    const pendingBefore = {
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    };
    const pendingAfter = {
      pendingStylisticArtifacts: 1,
      hasPendingStylisticArtifacts: true,
      trackChangesActive: true,
    };
    const factory = new ResolveSuggestionResultFactory("accept", {
      deriveDocumentState: vi.fn(
        (_reviewState: DocumentReviewState): DocumentReviewUiState =>
          "pending-review",
      ),
      inspect: vi.fn().mockResolvedValue(pendingAfter),
    });

    const result = await factory.buildObservationFailureResult(
      {} as Word.RequestContext,
      "identity-lost",
      pendingBefore,
    );

    expect(result.status).toBe("identity-lost");
    expect(result.error).toContain("compound-v2");
    expect(result.pendingAfter).toBe(pendingAfter);
  });

  it("builds a stable outer-catch error result from the provided fallback state", () => {
    const pendingAfter = {
      pendingStylisticArtifacts: 0,
      hasPendingStylisticArtifacts: false,
      trackChangesActive: false,
    };
    const deriveDocumentState = vi.fn(
      (_reviewState: DocumentReviewState): DocumentReviewUiState => "idle",
    );
    const factory = new ResolveSuggestionResultFactory("reject", {
      deriveDocumentState,
      inspect: vi.fn(),
    });

    const result = factory.buildErrorResult("boom", pendingAfter);

    expect(result).toEqual({
      status: "error",
      trackedChangesAffected: 0,
      commentDeleted: false,
      pendingAfter,
      documentState: "idle",
      error: "boom",
    });
    expect(deriveDocumentState).toHaveBeenCalledWith(pendingAfter);
  });
});
