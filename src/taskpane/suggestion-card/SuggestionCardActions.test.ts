import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { SuggestionResolutionMediatorResult } from "../../domain/suggestion/SuggestionResolutionWorkflow.types";
import { SuggestionStateMachine } from "../../domain/suggestion/SuggestionStateMachine";
import { createSuggestionProgressSummaryModel } from "../SuggestionProgressSummary";
import {
  createTaskpaneDocument,
  makeSuggestion,
  resetTaskpaneHarness,
  teardownTaskpaneHarness,
} from "../TaskpaneTestHelper";
import type { ResultsPanelDeps } from "../SuggestionCardRenderer.types";
import { createSuggestionCard } from "./SuggestionCardElements";
import { handleAcceptSuggestion } from "./SuggestionCardActions";

/** Builds a compact accepted mediator result for card-action tests. */
function makeAcceptedResult(): SuggestionResolutionMediatorResult {
  return {
    status: "accepted",
    trackedChangesAffected: 2,
    commentDeleted: true,
    pendingAfter: {
      pendingStylisticArtifacts: 0,
      hasPendingStylisticArtifacts: false,
      trackChangesActive: true,
    },
    documentState: "ready-to-disable-track-changes",
    feedbackStatus: "sent",
    taskpaneState: {
      documentState: "ready-to-disable-track-changes",
      showDisableTrackChangesCta: true,
      showCleanupSection: true,
    },
  };
}

describe("SuggestionCardActions", () => {
  beforeEach(() => {
    resetTaskpaneHarness();
    globalThis.document = createTaskpaneDocument() as unknown as Document;
  });

  afterEach(() => {
    teardownTaskpaneHarness();
  });

  it("accepts with the card feedback comment and applies terminal UI consequences", async () => {
    const suggestion = makeSuggestion({ id: "s-action" });
    const card = createSuggestionCard(suggestion, []);
    const textarea = card.li.querySelector(".feedback-textarea") as
      | HTMLTextAreaElement
      | null;
    const acceptBtn = card.li.querySelector(
      '[data-action="accept"]',
    ) as HTMLButtonElement | null;
    const rejectBtn = card.li.querySelector(
      '[data-action="reject"]',
    ) as HTMLButtonElement | null;
    const summaryElement = document.createElement("div");
    const deps: ResultsPanelDeps = {
      navigateToText: vi.fn(),
      acceptSuggestion: vi.fn().mockResolvedValue(makeAcceptedResult()),
      rejectSuggestion: vi.fn(),
    };

    if (!textarea) throw new Error("Missing feedback textarea");
    textarea.value = "  comentario útil  ";

    await handleAcceptSuggestion(
      suggestion,
      card.li,
      { acceptBtn, rejectBtn },
      new SuggestionStateMachine(),
      deps,
      {
        summaryModel: createSuggestionProgressSummaryModel(
          [suggestion],
          {
            successCount: 1,
            failedSuggestions: [],
            pendingAfter: {
              pendingStylisticArtifacts: 1,
              hasPendingStylisticArtifacts: true,
              trackChangesActive: true,
            },
            documentState: "pending-review",
            trackChangesActivatedForBatch: true,
          },
          [],
        ),
        summaryElement,
        isSelection: false,
      },
    );

    expect(deps.acceptSuggestion).toHaveBeenCalledWith(
      suggestion,
      "comentario útil",
    );
    expect(card.li.classList.contains("result-accepted")).toBe(true);
    expect(card.li.querySelector(".result-actions")).toBeNull();
    expect(document.getElementById("cleanup-section")?.style.display).toBe(
      "block",
    );
    expect(
      document.getElementById("disable-track-changes-section")?.style.display,
    ).toBe("block");
    expect(summaryElement.textContent).toContain(
      "Ya no te quedan sugerencias aplicadas por revisar.",
    );
    expect(summaryElement.textContent).toContain("1 ya resuelta.");
  });
});
