import { describe, expect, it } from "vitest";
import {
  DocumentReviewStateMachine,
  InvalidDocumentReviewTransitionError,
} from "./DocumentReviewStateMachine";
import type { DocumentReviewState } from "../types";

function makeReviewState(
  overrides: Partial<DocumentReviewState> = {},
): DocumentReviewState {
  return {
    pendingStylisticArtifacts: 0,
    hasPendingStylisticArtifacts: false,
    trackChangesActive: false,
    ...overrides,
  };
}

describe("DocumentReviewStateMachine", () => {
  it("derives pending-review when pending Stylistic artifacts exist", () => {
    const machine = new DocumentReviewStateMachine(
      makeReviewState({
        pendingStylisticArtifacts: 2,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      }),
    );

    expect(machine.state).toBe("pending-review");
    expect(machine.shouldShowDisableTrackChangesCta).toBe(false);
  });

  it("derives ready-to-disable-track-changes when zero pending remain but Track Changes stays active", () => {
    const machine = new DocumentReviewStateMachine(
      makeReviewState({ trackChangesActive: true }),
    );

    expect(machine.state).toBe("ready-to-disable-track-changes");
    expect(machine.shouldShowDisableTrackChangesCta).toBe(true);
  });

  it("evaluates the zero-pending transition and exposes the CTA semantic", () => {
    const transition = DocumentReviewStateMachine.evaluateTransition(
      makeReviewState({
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      }),
      makeReviewState({ trackChangesActive: true }),
    );

    expect(transition).toEqual({
      from: "pending-review",
      to: "ready-to-disable-track-changes",
    });
  });

  it("does not expose the CTA when Track Changes is already off after zero pending", () => {
    const transition = DocumentReviewStateMachine.evaluateTransition(
      makeReviewState({
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      }),
      makeReviewState({ trackChangesActive: false }),
    );

    expect(transition).toEqual({
      from: "pending-review",
      to: "idle",
    });
  });

  it("throws when disabling Track Changes from a non-ready state", () => {
    const machine = new DocumentReviewStateMachine(
      makeReviewState({
        pendingStylisticArtifacts: 1,
        hasPendingStylisticArtifacts: true,
        trackChangesActive: true,
      }),
    );

    expect(() => machine.disableTrackChanges()).toThrow(
      InvalidDocumentReviewTransitionError,
    );
  });
});
