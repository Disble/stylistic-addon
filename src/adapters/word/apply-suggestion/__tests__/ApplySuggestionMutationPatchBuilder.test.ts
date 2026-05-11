import { describe, expect, it } from "vitest";
import { makeSuggestion } from "../../ApplySuggestionCommandTestHelper";
import { ApplySuggestionMutationPatchBuilder } from "../ApplySuggestionMutationPatchBuilder";

describe("ApplySuggestionMutationPatchBuilder", () => {
  it("classifies delete, insert, and replace tracked-change mutations", () => {
    const builder = new ApplySuggestionMutationPatchBuilder();

    expect(builder.classifyChange(makeSuggestion({ suggestedText: "" }))).toBe("delete");
    expect(
      builder.classifyChange(
        makeSuggestion({ anchor: "", context: "", suggestedText: "nuevo texto" })
      )
    ).toBe("insert");
    expect(builder.classifyChange(makeSuggestion())).toBe("replace");
  });

  it("builds a localized mutation patch from the container text", () => {
    const builder = new ApplySuggestionMutationPatchBuilder();

    expect(
      builder.buildApplyMutationPatch(
        makeSuggestion({
          id: "patch-1",
          anchor: "texto original",
          suggestedText: "texto sugerido",
          positionHint: {
            start: 10,
            end: 24,
            snapshotVersion: 4,
            paragraphId: "p-1",
            source: "snapshot",
          },
        }),
        "Antes texto original y después."
      )
    ).toEqual({
      suggestionId: "patch-1",
      snapshotVersion: 5,
      paragraphId: "p-1",
      originalText: "Antes texto original y después.",
      updatedText: "Antes texto sugerido y después.",
      deltaLength: 0,
      affectedStart: 6,
      affectedEnd: 20,
    });
  });

  it("stringifies unknown errors without throwing", () => {
    const builder = new ApplySuggestionMutationPatchBuilder();

    expect(builder.stringifyUnknownError(new Error("boom"))).toBe("boom");
    expect(builder.stringifyUnknownError("boom")).toBe("boom");
    expect(builder.stringifyUnknownError({ reason: "boom" })).toBe('{"reason":"boom"}');
  });
});
