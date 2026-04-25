import { describe, expect, it, vi } from "vitest";
import type { Suggestion } from "../../../domain/types";
import { SuggestionResolutionObserver } from "./SuggestionResolutionObserver";
import type { ReplaceObservationContext } from "./ResolutionContext";

type ObserverTestDouble = {
  observeResolutionCandidate: (
    context: Word.RequestContext,
    cc: Word.ContentControl,
  ) => Promise<ReplaceObservationContext>;
};

/** Builds one minimal track-change suggestion for observer-focused regressions. */
function makeTrackChangeSuggestion(): Suggestion {
  return {
    id: "chunk0-0",
    context: "No sabían si venía de ni Shu o de otro sitio.",
    anchor: "ni Shu",
    suggestedText: "ni de Shu",
    justification: "Mas claro",
    category: "Claridad",
    severity: "medium",
    type: "track-change",
  };
}

describe("SuggestionResolutionObserver.observeResolutionCandidates", () => {
  it("keeps the last observed tracked-change evidence when a later candidate throws", async () => {
    const suggestion = makeTrackChangeSuggestion();
    const locator = {
      findColocatedStylisticComment: vi.fn().mockResolvedValue(null),
    };
    const observer = new SuggestionResolutionObserver(
      suggestion,
      locator as never,
      {} as never,
    );

    const firstCandidate = {
      tag: "stylistic:track-change:chunk0-0",
    } as Word.ContentControl;
    const secondCandidate = {
      tag: "stylistic:track-change:chunk0-0-shadow",
    } as Word.ContentControl;
    const deletedTrackedChange = {
      type: "Deleted",
    } as Word.TrackedChange;
    const firstObservation: ReplaceObservationContext = {
      trackedChanges: [deletedTrackedChange],
      observationStatus: "unobservable",
      semanticCandidates: {
        Deleted: [
          {
            trackedChange: deletedTrackedChange,
            source: "cc",
          },
        ],
        Added: [],
      },
    };

    const observeResolutionCandidateSpy = vi.spyOn(
      observer as unknown as ObserverTestDouble,
      "observeResolutionCandidate",
    );
    observeResolutionCandidateSpy.mockResolvedValueOnce(firstObservation);
    observeResolutionCandidateSpy.mockRejectedValueOnce(
      new Error("GeneralException"),
    );

    const result = await observer.observeResolutionCandidates(
      {} as Word.RequestContext,
      [firstCandidate, secondCandidate],
      firstCandidate,
    );

    expect(result.selectedCc).toBe(firstCandidate);
    expect(result.trackedChanges).toEqual([deletedTrackedChange]);
    expect(result.observationStatus).toBe("unobservable");
    expect(result.semanticCandidates?.Deleted).toHaveLength(1);
    expect(locator.findColocatedStylisticComment).toHaveBeenCalledTimes(2);
  });

  it("still propagates the failure when no prior candidate produced evidence", async () => {
    const suggestion = makeTrackChangeSuggestion();
    const locator = {
      findColocatedStylisticComment: vi.fn().mockResolvedValue(null),
    };
    const observer = new SuggestionResolutionObserver(
      suggestion,
      locator as never,
      {} as never,
    );

    const candidate = {
      tag: "stylistic:track-change:chunk0-0",
    } as Word.ContentControl;

    const observeResolutionCandidateSpy = vi.spyOn(
      observer as unknown as ObserverTestDouble,
      "observeResolutionCandidate",
    );
    observeResolutionCandidateSpy.mockRejectedValueOnce(
      new Error("GeneralException"),
    );

    await expect(
      observer.observeResolutionCandidates(
        {} as Word.RequestContext,
        [candidate],
        candidate,
      ),
    ).rejects.toThrow("GeneralException");
  });
});
