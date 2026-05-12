import { describe, expect, it, vi } from "vitest";
import type { Suggestion } from "../../../../domain/suggestion/Suggestion.types";
import {
  makeOperationalWrapperTag,
  makeOperationalWrapperTitle,
  makeResolveSuggestionContext,
} from "../../WordAdapterActionTestHelper";
import type { TextLocator } from "../../WordTextLocatorContext.types";
import { SuggestionLocator } from "../SuggestionLocator";
import { SuggestionResolutionObserver } from "../SuggestionResolutionObserver";

const NOOP_TEXT_LOCATOR: TextLocator = {
  locate: vi.fn(async () => null),
};

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
  function castToObservationArgs(context: ReturnType<typeof makeResolveSuggestionContext>): {
    requestContext: Word.RequestContext;
    candidates: Word.ContentControl[];
    selectedCc: Word.ContentControl;
  } {
    return {
      requestContext: context as unknown as Word.RequestContext,
      candidates: context._ccItems as unknown as Word.ContentControl[],
      selectedCc: context._cc as unknown as Word.ContentControl,
    };
  }

  it("observes the wrapper range collection as the executable operational scope", async () => {
    const suggestion = makeTrackChangeSuggestion();
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeOperationalWrapperTag("chunk0-0"),
      ccTitle: makeOperationalWrapperTitle({
        suggestionId: "chunk0-0",
        insertedTag: "stylistic:track-change:chunk0-0",
        deletedValue: "ni Shu",
        anchorValue: "No sabían si venía de ni Shu o de otro sitio.",
      }),
      rangeTCItems: [
        { id: "tc-added", type: "Added", accept: vi.fn(), reject: vi.fn() },
        { id: "tc-deleted", type: "Deleted", accept: vi.fn(), reject: vi.fn() },
      ],
      comments: [],
    });
    const locator = new SuggestionLocator(suggestion);
    const observer = new SuggestionResolutionObserver(suggestion, locator, NOOP_TEXT_LOCATOR);

    const args = castToObservationArgs(context);
    const result = await observer.observeResolutionCandidates(
      args.requestContext,
      args.candidates,
      args.selectedCc
    );

    expect(result.observationStatus).toBe("confirmed-pending");
    expect(result.trackedChanges).toHaveLength(2);
    expect(result.trackedChangesCollection).toBeDefined();
  });

  it("fails closed as mixed-group when the wrapper belongs to a contiguous group", async () => {
    const suggestion = makeTrackChangeSuggestion();
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeOperationalWrapperTag("chunk0-0"),
      ccItems: [
        {
          tag: makeOperationalWrapperTag("chunk0-0"),
          title: makeOperationalWrapperTitle({
            suggestionId: "chunk0-0",
            insertedTag: "stylistic:track-change:chunk0-0",
            deletedValue: "ni Shu",
            anchorValue: "No sabían si venía de ni Shu o de otro sitio.",
            groupId: "group-a",
            groupIndex: 0,
            groupSize: 2,
          }),
          rangeRelationWithNext: "AdjacentBefore",
        },
        {
          tag: makeOperationalWrapperTag("chunk0-1"),
          title: makeOperationalWrapperTitle({
            suggestionId: "chunk0-1",
            insertedTag: "stylistic:track-change:chunk0-1",
            deletedValue: "otro",
            anchorValue: "Otro contexto.",
            groupId: "group-a",
            groupIndex: 1,
            groupSize: 2,
          }),
        },
      ],
      comments: [],
    });
    const locator = new SuggestionLocator(suggestion);
    const observer = new SuggestionResolutionObserver(suggestion, locator, NOOP_TEXT_LOCATOR);

    const selectedCc = context._ccItems[0];
    expect(selectedCc).toBeDefined();
    if (!selectedCc) {
      throw new Error("Expected first operational wrapper candidate.");
    }

    const result = await observer.observeResolutionCandidates(
      context as unknown as Word.RequestContext,
      context._ccItems as unknown as Word.ContentControl[],
      selectedCc as unknown as Word.ContentControl
    );

    expect(result.observationStatus).toBe("mixed-group");
    expect(result.trackedChanges).toEqual([]);
  });

  it("accepts delete-only wrapper identities and treats wrapper tracked changes as pending evidence", async () => {
    const suggestion = {
      ...makeTrackChangeSuggestion(),
      id: "delete-only-1",
      anchor: " a pesar de eso",
      context: "No obstante, siguió sosteniéndola del brazo a pesar de eso.",
      suggestedText: "",
    };
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeOperationalWrapperTag("delete-only-1"),
      ccTitle: makeOperationalWrapperTitle({
        suggestionId: "delete-only-1",
        trackChangeSubtype: "delete-only",
        deleteValue: " a pesar de eso",
        anchorValue: "No obstante, siguió sosteniéndola del brazo a pesar de eso.",
      }),
      rangeTCItems: [
        { id: "tc-delete", type: "Deleted", accept: vi.fn(), reject: vi.fn() },
        { id: "tc-empty", type: "Added", accept: vi.fn(), reject: vi.fn() },
      ],
      comments: [],
    });
    const locator = new SuggestionLocator(suggestion);
    const observer = new SuggestionResolutionObserver(suggestion, locator, NOOP_TEXT_LOCATOR);

    const args = castToObservationArgs(context);
    const result = await observer.observeResolutionCandidates(
      args.requestContext,
      args.candidates,
      args.selectedCc
    );

    expect(result.observationStatus).toBe("confirmed-pending");
    expect(result.trackedChanges).toHaveLength(2);
  });

  it("accepts formatting wrapper identities and treats Formatted changes as pending evidence", async () => {
    const suggestion = {
      ...makeTrackChangeSuggestion(),
      id: "format-1",
      anchor: "post mortem",
      context: "Ese era el inicio del post mortem reportado por PRIME.",
      suggestedText: "*post mortem*",
    };
    const context = makeResolveSuggestionContext({
      ccFound: true,
      ccTag: makeOperationalWrapperTag("format-1"),
      ccTitle: makeOperationalWrapperTitle({
        suggestionId: "format-1",
        trackChangeSubtype: "formatting",
        formatTag: "stylistic:track-change:format-1",
        anchorValue: "Ese era el inicio del post mortem reportado por PRIME.",
      }),
      rangeTCItems: [{ id: "tc-format", type: "Formatted", accept: vi.fn(), reject: vi.fn() }],
      comments: [],
    });
    const locator = new SuggestionLocator(suggestion);
    const observer = new SuggestionResolutionObserver(suggestion, locator, NOOP_TEXT_LOCATOR);

    const args = castToObservationArgs(context);
    const result = await observer.observeResolutionCandidates(
      args.requestContext,
      args.candidates,
      args.selectedCc
    );

    expect(result.observationStatus).toBe("confirmed-pending");
    expect(result.trackedChanges).toHaveLength(1);
  });
});
